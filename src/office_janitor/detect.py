"""!
@brief Detection helpers for MSI and Click-to-Run Office deployments.
@details Reads structured metadata from :mod:`office_janitor.constants`, probes
registry hives, and returns structured :class:`DetectedInstallation` records that
contain uninstall handles, source type, and channel information.
"""

from __future__ import annotations

import csv
import json
import logging
import os
import shlex
import sys
from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from . import constants, elevation, exec_utils, logging_ext, registry_tools

_LOGGER = logging.getLogger(__name__)


_OFFICE_PROCESS_TARGETS = tuple(
    sorted(
        {name.lower() for name in constants.DEFAULT_OFFICE_PROCESSES} | {"mspub.exe", "teams.exe"}
    )
)
"""!
@brief Known Office executables monitored during detection.
"""

_SERVICE_TARGETS = tuple(sorted({name.lower() for name in constants.KNOWN_SERVICES} | {"osppsvc"}))
"""!
@brief Services associated with Office provisioning and licensing.
"""

_TASK_PREFIXES = (r"\\Microsoft\\Office\\", r"\\Microsoft\\OfficeSoftwareProtectionPlatform\\")
"""!
@brief Scheduled task prefixes that indicate Office automation jobs.
"""

_KNOWN_TASK_NAMES = {
    task if task.startswith("\\") else f"\\{task}" for task in constants.KNOWN_SCHEDULED_TASKS
}
"""!
@brief Explicit scheduled task identifiers from the specification.
"""


@dataclass(frozen=True)
class DetectedInstallation:
    """!
    @brief Structured record describing an installed Office product.
    @details Provides a serialisable view of the data gathered from registry
    hives so downstream planners can differentiate MSI and Click-to-Run
    inventory without re-querying the registry.
    """

    source: str
    product: str
    version: str
    architecture: str
    uninstall_handles: tuple[str, ...]
    channel: str
    product_code: str | None = None
    release_ids: tuple[str, ...] = ()
    properties: Mapping[str, object] | None = None
    display_icon: str | None = None
    maintenance_paths: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, object]:
        """!
        @brief Convert the dataclass to a JSON-serialisable dictionary.
        """

        payload: dict[str, object] = {
            "source": self.source,
            "product": self.product,
            "version": self.version,
            "architecture": self.architecture,
            "uninstall_handles": list(self.uninstall_handles),
            "channel": self.channel,
        }
        if self.product_code:
            payload["product_code"] = self.product_code
        if self.release_ids:
            payload["release_ids"] = list(self.release_ids)
        if self.properties:
            payload["properties"] = dict(self.properties)
        if self.display_icon:
            payload["display_icon"] = self.display_icon
        if self.maintenance_paths:
            payload["maintenance_paths"] = list(self.maintenance_paths)
        return payload


def _strip_icon_index(raw: str) -> str:
    """!
    @brief Remove trailing icon index fragments from registry path values.
    """

    text = raw.strip()
    if "," not in text:
        return text
    prefix, _, suffix = text.partition(",")
    remainder = suffix.strip()
    if not remainder:
        return prefix
    if remainder.lstrip("+-").isdigit():
        return prefix
    return text


def _extract_executable_candidate(value: object) -> str:
    """!
    @brief Attempt to extract an executable path from ``value``.
    """

    if value is None:
        return ""
    text = _strip_icon_index(str(value).strip())
    if not text:
        return ""
    cleaned_text = text.strip().strip('"').strip()
    candidates: list[str] = []
    if cleaned_text:
        candidates.append(cleaned_text)
    try:
        parts = shlex.split(text, posix=False)
    except ValueError:
        parts = []
    if parts:
        candidates.extend(parts)
    for token in candidates:
        cleaned = token.strip().strip('"').strip()
        if not cleaned:
            continue
        lowered = cleaned.lower()
        if lowered.endswith(".exe"):
            return cleaned
        if "setup.exe" in lowered:
            index = lowered.find("setup.exe")
            return cleaned[: index + len("setup.exe")]
    return ""


def _collect_maintenance_paths(values: Mapping[str, object]) -> tuple[str, ...]:
    """!
    @brief Collect setup.exe maintenance candidates from registry values.
    """

    candidates: list[str] = []
    for key in ("DisplayIcon", "ModifyPath", "UninstallString"):
        candidate = _extract_executable_candidate(values.get(key))
        if candidate and candidate.lower().endswith("setup.exe"):
            if candidate not in candidates:
                candidates.append(candidate)
    return tuple(candidates)


def _friendly_channel(raw_channel: str | None) -> str:
    """!
    @brief Resolve a friendly channel name for Click-to-Run metadata.
    """

    if not raw_channel:
        return "unknown"
    return constants.C2R_CHANNEL_ALIASES.get(raw_channel, raw_channel)


def _compose_handle(root: int, path: str) -> str:
    """!
    @brief Helper to create a ``HKLM\\...`` style registry handle identifier.
    """

    return f"{registry_tools.hive_name(root)}\\{path}"


def _safe_read_values(root: int, path: str) -> dict[str, Any]:
    """!
    @brief Read registry values while tolerating missing hives or permissions.
    @details Wraps :func:`registry_tools.read_values` so callers can safely probe
    both 32-bit and 64-bit views without surfacing platform-specific
    exceptions during detection runs on non-Windows hosts.
    """

    try:
        return registry_tools.read_values(root, path)
    except FileNotFoundError:
        return {}
    except OSError:
        return {}


def _powershell_escape(text: str) -> str:
    """!
    @brief Escape a string literal for embedding within PowerShell commands.
    """

    return text.replace("'", "''")


def _powershell_registry_path(root: int, path: str) -> str:
    """!
    @brief Convert registry handles into PowerShell provider paths.
    """

    hive = registry_tools.hive_name(root)
    normalized = path.replace("/", "\\")
    return f"{hive}:\\{normalized}"


def _powershell_read_values(root: int, path: str) -> dict[str, Any]:
    """!
    @brief Query registry values using ``powershell`` as a fallback.
    @details ``winreg`` is unavailable on some unit test hosts. This helper
    mirrors the value projection from :func:`registry_tools.read_values` by
    invoking PowerShell to emit JSON for the requested key. Errors return an
    empty mapping so detection can degrade gracefully.
    """

    provider_path = _powershell_registry_path(root, path)
    script = (
        "$ErrorActionPreference='SilentlyContinue';"
        f"$p='{_powershell_escape(provider_path)}';"
        "if(Test-Path $p){"
        "  Get-ItemProperty -Path $p | Select-Object * | ConvertTo-Json -Compress"
        "}else{''}"
    )
    code, output = _run_command(["powershell", "-NoProfile", "-Command", script])
    if code != 0 or not output.strip():
        return {}

    text = output.strip()
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        return {}

    if isinstance(data, list):
        data = data[0] if data else {}
    if not isinstance(data, dict):
        return {}

    filtered: dict[str, Any] = {}
    for key, value in data.items():
        if str(key).startswith("PS"):
            continue
        filtered[str(key)] = value
    return filtered


def _read_values_with_fallback(root: int, path: str) -> dict[str, Any]:
    """!
    @brief Read registry values with a PowerShell fallback when necessary.
    """

    values = _safe_read_values(root, path)
    if values:
        return values
    return _powershell_read_values(root, path)


def _key_exists_with_fallback(root: int, path: str) -> bool:
    """!
    @brief Determine whether ``root\\path`` exists using registry or PowerShell probes.
    """

    try:
        if registry_tools.key_exists(root, path):
            return True
    except FileNotFoundError:
        pass
    except OSError:
        pass

    provider_path = _powershell_registry_path(root, path)
    script = (
        "$ErrorActionPreference='SilentlyContinue';"
        f"if(Test-Path '{_powershell_escape(provider_path)}'){{'True'}}else{{'False'}}"
    )
    code, output = _run_command(["powershell", "-NoProfile", "-Command", script])
    if code != 0 or not output.strip():
        return False
    result = output.strip().splitlines()[-1].strip().lower()
    return result == "true"


def _parse_languages(*candidates: object) -> tuple[str, ...]:
    """!
    @brief Normalise language identifiers from registry and subscription data.
    """

    languages: list[str] = []

    def add_token(token: str) -> None:
        cleaned = token.strip()
        if not cleaned:
            return
        lower = cleaned.lower()
        for existing in languages:
            if existing.lower() == lower:
                return
        languages.append(cleaned)

    for candidate in candidates:
        if not candidate:
            continue
        if isinstance(candidate, str):
            expanded = candidate.replace(";", ",").replace("|", ",")
            for part in expanded.split(","):
                add_token(part)
            continue
        if isinstance(candidate, (list, tuple, set)):
            for part in candidate:
                if isinstance(part, str):
                    add_token(part)
        elif isinstance(candidate, Mapping):
            for part in candidate.values():
                if isinstance(part, str):
                    add_token(part)

    return tuple(languages)


def _infer_architecture(name: str, install_path: str | None = None) -> str:
    """!
    @brief Infer architecture information from names or installation paths.
    """

    lowered = name.lower()
    if "64-bit" in lowered or "(64" in lowered or "x64" in lowered:
        return "x64"
    if "32-bit" in lowered or "(32" in lowered or "x86" in lowered:
        return "x86"
    if install_path:
        path_lower = install_path.lower()
        if "program files (x86)" in path_lower:
            return "x86"
        if "program files" in path_lower:
            return "x64"
    return "unknown"


def _merge_fallback_metadata(
    existing: dict[str, dict[str, Any]], additions: Mapping[str, Mapping[str, Any]]
) -> None:
    """!
    @brief Merge fallback detection metadata keyed by product code.
    """

    for code, payload in additions.items():
        if not code:
            continue
        key = code.upper()
        current = existing.get(key)
        if current is None:
            existing[key] = dict(payload)
            continue
        merged = dict(current)
        merged.update(payload)
        existing[key] = merged


def _candidate_msi_handles(product_code: str) -> tuple[str, ...]:
    """!
    @brief Determine candidate uninstall handles for an MSI product code.
    """

    handles: list[str] = []
    for hive, base_key in constants.MSI_UNINSTALL_ROOTS:
        key_path = f"{base_key}\\{product_code}"
        if _key_exists_with_fallback(hive, key_path):
            handles.append(_compose_handle(hive, key_path))
    if not handles and constants.MSI_UNINSTALL_ROOTS:
        hive, base_key = constants.MSI_UNINSTALL_ROOTS[0]
        handles.append(_compose_handle(hive, f"{base_key}\\{product_code}"))
    return tuple(handles)


def _read_subscription_values(release_id: str) -> tuple[dict[str, Any], str | None]:
    """!
    @brief Read Click-to-Run subscription metadata for ``release_id``.
    """

    for hive, base_path in constants.C2R_SUBSCRIPTION_ROOTS:
        key_path = f"{base_path}\\{release_id}"
        values = _read_values_with_fallback(hive, key_path)
        if values:
            return values, _compose_handle(hive, key_path)
    return {}, None


def _normalize_release_ids(raw: object) -> tuple[str, ...]:
    """!
    @brief Convert registry values into a canonical tuple of release identifiers.
    """

    tokens: list[str] = []
    if isinstance(raw, str):
        expanded = raw.replace(";", ",").replace("|", ",")
        tokens = [segment.strip() for segment in expanded.split(",") if segment.strip()]
    elif isinstance(raw, (list, tuple, set)):
        for entry in raw:
            text = str(entry).strip()
            if text:
                tokens.append(text)

    seen: set[str] = set()
    ordered: list[str] = []
    for token in tokens:
        normalized = token
        if normalized in seen:
            continue
        seen.add(normalized)
        ordered.append(normalized)
    return tuple(ordered)


def _probe_msi_wmi() -> dict[str, dict[str, Any]]:
    """!
    @brief Collect MSI product metadata via ``wmic`` when available.
    """

    code, output = _run_command(
        ["wmic", "product", "get", "IdentifyingNumber,Name,Version,InstallLocation", "/format:csv"]
    )
    if code != 0 or not output.strip():
        return {}

    lines = [line.strip("\ufeff").strip() for line in output.splitlines() if line.strip()]
    if not lines:
        return {}

    try:
        reader = csv.DictReader(lines)
    except csv.Error:
        return {}

    results: dict[str, dict[str, Any]] = {}
    for row in reader:
        product_code = str(row.get("IdentifyingNumber") or "").strip()
        name = str(row.get("Name") or "").strip()
        version = str(row.get("Version") or "").strip()
        install_path = str(row.get("InstallLocation") or "").strip()
        if not product_code and not name:
            continue
        if name and not any(token in name.lower() for token in ("office", "visio", "project")):
            continue
        results[product_code.upper()] = {
            "product": name or product_code,
            "version": version,
            "install_location": install_path,
            "probe": "wmic",
        }

    return results


def _probe_msi_powershell() -> dict[str, dict[str, Any]]:
    """!
    @brief Collect MSI product metadata via ``powershell`` CIM queries when available.
    """

    script = (
        "$ErrorActionPreference='SilentlyContinue';"
        "$items=Get-CimInstance -ClassName Win32_Product -ErrorAction SilentlyContinue | "
        "Where-Object { $_.Name -and ( "
        "$_.Name -like '*Office*' -or $_.Name -like '*Visio*' -or $_.Name -like '*Project*' ) } | "
        "Select-Object IdentifyingNumber,Name,Version,InstallLocation;"
        "if($items){$items|ConvertTo-Json -Compress}else{''}"
    )
    code, output = _run_command(["powershell", "-NoProfile", "-Command", script])
    if code != 0 or not output.strip():
        return {}

    text = output.strip()
    try:
        payload = json.loads(text)
    except json.JSONDecodeError:
        return {}

    records = payload if isinstance(payload, list) else [payload]
    results: dict[str, dict[str, Any]] = {}
    for record in records:
        if not isinstance(record, Mapping):
            continue
        product_code = str(record.get("IdentifyingNumber") or "").strip()
        name = str(record.get("Name") or "").strip()
        version = str(record.get("Version") or "").strip()
        install_path = str(record.get("InstallLocation") or "").strip()
        if not product_code and not name:
            continue
        results[product_code.upper()] = {
            "product": name or product_code,
            "version": version,
            "install_location": install_path,
            "probe": "powershell",
        }

    return results


def detect_msi_installations() -> list[DetectedInstallation]:
    """!
    @brief Inspect the registry and return metadata for MSI-based Office installs.
    """

    installations: list[DetectedInstallation] = []
    seen_handles: set[str] = set()
    seen_codes: set[str] = set()

    fallback_sources: dict[str, dict[str, Any]] = {}
    _merge_fallback_metadata(fallback_sources, _probe_msi_wmi())
    _merge_fallback_metadata(fallback_sources, _probe_msi_powershell())

    for product_code, metadata in constants.MSI_PRODUCT_MAP.items():
        registry_roots: Iterable[tuple[int, str]] = metadata.get(
            "registry_roots", constants.MSI_UNINSTALL_ROOTS
        )
        for hive, base_key in registry_roots:
            key_path = f"{base_key}\\{product_code}"
            values = _read_values_with_fallback(hive, key_path)
            if not values:
                continue

            handle = _compose_handle(hive, key_path)
            if handle in seen_handles:
                continue

            fallback_meta = fallback_sources.pop(product_code.upper(), None)
            display_name = str(values.get("DisplayName") or metadata.get("product") or product_code)
            display_version = str(
                values.get("DisplayVersion")
                or metadata.get("version")
                or (fallback_meta or {}).get("version")
                or "unknown"
            )
            uninstall_string = str(values.get("UninstallString") or "")
            install_location = str(
                values.get("InstallLocation") or (fallback_meta or {}).get("install_location") or ""
            )
            display_icon_value = str(values.get("DisplayIcon") or "").strip()
            maintenance_paths = _collect_maintenance_paths(values)
            raw_architecture = str(metadata.get("architecture", "")).strip()
            architecture = raw_architecture or _infer_architecture(
                display_name, install_location or None
            )
            if not architecture:
                architecture = "unknown"
            family = constants.resolve_msi_family(product_code) or str(metadata.get("family", ""))
            languages = _parse_languages(
                values.get("InstallLanguage"),
                values.get("Language"),
                values.get("ProductLanguage"),
            )

            properties: dict[str, object] = {
                "display_name": display_name,
                "display_version": display_version,
                "supported_versions": list(metadata.get("supported_versions", ())),
                "edition": metadata.get("edition", ""),
            }
            if uninstall_string:
                properties["uninstall_string"] = uninstall_string
            if install_location:
                properties["install_location"] = install_location
            if display_icon_value:
                properties["display_icon"] = display_icon_value
            if maintenance_paths:
                properties["maintenance_paths"] = list(maintenance_paths)
            if family:
                properties["family"] = family
            if languages:
                properties["languages"] = list(languages)
            if fallback_meta and fallback_meta.get("probe"):
                properties["supplemental_probes"] = [str(fallback_meta["probe"])]

            installations.append(
                DetectedInstallation(
                    source="MSI",
                    product=str(metadata.get("product", display_name)),
                    version=display_version or "unknown",
                    architecture=architecture or "unknown",
                    uninstall_handles=(handle,),
                    channel="MSI",
                    product_code=product_code,
                    properties=properties,
                    display_icon=display_icon_value or None,
                    maintenance_paths=maintenance_paths,
                )
            )
            seen_handles.add(handle)
            seen_codes.add(product_code.upper())

    for product_code, metadata in fallback_sources.items():
        if product_code in seen_codes:
            continue
        display_name = str(metadata.get("product") or product_code)
        version = str(metadata.get("version") or "unknown") or "unknown"
        install_location = str(metadata.get("install_location") or "")
        architecture = _infer_architecture(display_name, install_location or None) or "unknown"
        handles = _candidate_msi_handles(product_code)
        properties: dict[str, object] = {
            "display_name": display_name,
            "display_version": version,
        }
        if install_location:
            properties["install_location"] = install_location
        probe = metadata.get("probe")
        if probe:
            properties["supplemental_probes"] = [str(probe)]

        installations.append(
            DetectedInstallation(
                source="MSI",
                product=display_name,
                version=version or "unknown",
                architecture=architecture,
                uninstall_handles=handles,
                channel="MSI",
                product_code=product_code,
                properties=properties,
                display_icon=None,
                maintenance_paths=(),
            )
        )

    return installations


def detect_c2r_installations() -> list[DetectedInstallation]:
    """!
    @brief Probe Click-to-Run configuration to describe installed suites.
    """

    installations: list[DetectedInstallation] = []

    for hive, config_path in constants.C2R_CONFIGURATION_KEYS:
        config_values = _read_values_with_fallback(hive, config_path)
        if not config_values:
            continue

        release_ids = _normalize_release_ids(
            config_values.get("ProductReleaseIds") or config_values.get("ReleaseIds")
        )
        if not release_ids:
            continue

        platform = str(
            config_values.get("Platform")
            or config_values.get("PlatformId")
            or config_values.get("PlatformCode")
            or ""
        ).lower()
        architecture = constants.C2R_PLATFORM_ALIASES.get(platform, platform or "unknown")
        version = str(
            config_values.get("VersionToReport")
            or config_values.get("ClientVersionToReport")
            or config_values.get("ProductVersion")
            or "unknown"
        )
        channel_identifier = (
            config_values.get("UpdateChannel")
            or config_values.get("ChannelId")
            or config_values.get("CDNBaseUrl")
        )
        global_channel = _friendly_channel(str(channel_identifier) if channel_identifier else None)
        package_guid = str(
            config_values.get("PackageGUID") or config_values.get("PackageGuid") or ""
        )
        install_path = str(
            config_values.get("InstallPath")
            or config_values.get("ClientFolder")
            or config_values.get("InstallOfficePath")
            or ""
        )
        configuration_languages = _parse_languages(
            config_values.get("Language"),
            config_values.get("InstallLanguage"),
            config_values.get("ClientCulture"),
            config_values.get("InstalledLanguages"),
        )
        base_handle = _compose_handle(hive, config_path)

        for release_id in release_ids:
            product_metadata = constants.C2R_PRODUCT_RELEASES.get(release_id)
            product_name = (
                str((product_metadata or {}).get("product", release_id))
                if product_metadata
                else release_id
            )
            supported_versions = tuple(
                str(v) for v in (product_metadata or {}).get("supported_versions", ())
            )
            supported_architectures = tuple(
                str(a) for a in (product_metadata or {}).get("architectures", ())
            )
            family = constants.resolve_c2r_family(release_id) or str(
                (product_metadata or {}).get("family", "")
            )

            uninstall_handles: set[str] = {base_handle}
            registry_paths = (product_metadata or {}).get("registry_paths", {})
            release_roots: Iterable[tuple[int, str]] = registry_paths.get(
                "product_release_ids", constants.C2R_PRODUCT_RELEASE_ROOTS
            )
            for rel_hive, rel_base in release_roots:
                release_key = f"{rel_base}\\{release_id}"
                if _key_exists_with_fallback(rel_hive, release_key):
                    uninstall_handles.add(_compose_handle(rel_hive, release_key))

            subscription_values, subscription_handle = _read_subscription_values(release_id)
            if subscription_handle:
                uninstall_handles.add(subscription_handle)
            subscription_channel = _friendly_channel(
                str(
                    subscription_values.get("ChannelId")
                    or subscription_values.get("UpdateChannel")
                    or subscription_values.get("CDNBaseUrl")
                    or ""
                )
                if subscription_values
                else None
            )
            release_channel = (
                subscription_channel if subscription_channel != "unknown" else global_channel
            )
            release_languages = _parse_languages(
                configuration_languages,
                subscription_values.get("Language"),
                subscription_values.get("InstalledLanguages"),
                subscription_values.get("ClientCulture"),
            )

            architecture_choice = architecture
            if architecture_choice == "unknown" and supported_architectures:
                architecture_choice = supported_architectures[0]

            properties: dict[str, object] = {
                "release_id": release_id,
                "version": version,
                "supported_versions": list(supported_versions),
                "supported_architectures": list(supported_architectures),
            }
            if package_guid:
                properties["package_guid"] = package_guid
            if install_path:
                properties["install_path"] = install_path
            if family:
                properties["family"] = family
            if release_languages:
                properties["languages"] = list(release_languages)
            if subscription_channel != "unknown":
                properties["channel_source"] = "subscription"
            else:
                properties["channel_source"] = "configuration"

            installations.append(
                DetectedInstallation(
                    source="C2R",
                    product=product_name,
                    version=version,
                    architecture=architecture_choice or "unknown",
                    uninstall_handles=tuple(sorted(uninstall_handles)),
                    channel=release_channel,
                    release_ids=(release_id,),
                    properties=properties,
                )
            )

    return installations


def gather_office_inventory(*, limited_user: bool | None = None) -> dict[str, object]:
    """!
    @brief Aggregate MSI, C2R, and ancillary signals into an inventory payload.
    """

    run_under_limited = bool(limited_user)
    human_logger = logging_ext.get_human_logger()
    if (
        run_under_limited
        and elevation.is_admin()
        and not os.environ.get("OFFICE_JANITOR_DEELEVATED")
    ):
        human_logger.info(
            "Detection requested under limited user context; attempting de-elevated probes."
        )
        result = elevation.run_as_limited_user(
            [sys.executable, "-m", "office_janitor.detect"],
            event="detect_deelevate",
            env_overrides={"OFFICE_JANITOR_DEELEVATED": "1"},
        )
        if result.returncode == 0 and result.stdout:
            try:
                parsed = json.loads(result.stdout)
                if isinstance(parsed, dict):
                    return parsed
            except json.JSONDecodeError:
                human_logger.warning(
                    "Failed to parse limited-user detection output; falling back to current "
                    "context."
                )

    inventory: dict[str, object] = {
        "context": {
            "user": elevation.current_username(),
            "is_admin": elevation.is_admin(),
        },
        "msi": [entry.to_dict() for entry in detect_msi_installations()],
        "c2r": [entry.to_dict() for entry in detect_c2r_installations()],
        "filesystem": [],
        "processes": gather_running_office_processes(),
        "services": gather_office_services(),
        "tasks": gather_office_tasks(),
        "activation": gather_activation_state(),
        "registry": gather_registry_residue(),
    }

    seen_paths: set[str] = set()

    for template in constants.INSTALL_ROOT_TEMPLATES:
        candidate = Path(template["path"])
        try:
            exists = candidate.exists()
        except OSError:
            exists = False
        if not exists:
            continue
        path_str = str(candidate)
        if path_str in seen_paths:
            continue
        inventory["filesystem"].append(
            {
                "path": str(candidate),
                "architecture": template.get("architecture", "unknown"),
                "release": template.get("release", ""),
                "label": template.get("label", ""),
            }
        )
        seen_paths.add(path_str)

    for template in constants.RESIDUE_PATH_TEMPLATES:
        raw_path = str(template.get("path", "").strip())
        if not raw_path:
            continue
        expanded = os.path.expandvars(raw_path)
        candidate = Path(expanded)
        try:
            exists = candidate.exists()
        except OSError:
            exists = False
        if not exists:
            continue
        path_str = str(candidate)
        if path_str in seen_paths:
            continue
        entry: dict[str, object] = {
            "path": path_str,
            "label": template.get("label", ""),
            "category": template.get("category", "residue"),
        }
        if "architecture" in template:
            entry["architecture"] = template["architecture"]
        inventory["filesystem"].append(entry)
        seen_paths.add(path_str)

    return inventory


def reprobe(options: Mapping[str, object] | None = None) -> dict[str, object]:
    """!
    @brief Re-run Office detection after a scrub pass to check for leftovers.
    @details The optional ``options`` mapping is accepted for parity with future
    targeted detection strategies, but currently serves only as a hook for
    logging and diagnostics. The returned inventory mirrors
    :func:`gather_office_inventory`.
    """

    _ = options  # Options are presently unused but reserved for parity.
    limited_user = None
    if isinstance(options, Mapping):
        limited_user = bool(options.get("limited_user"))
    return gather_office_inventory(limited_user=limited_user)


def _run_command(arguments: Iterable[str]) -> tuple[int, str]:
    """!
    @brief Execute a subprocess returning ``(returncode, text_output)``.
    @details Failures caused by missing binaries or platform limitations are
    normalised to a non-zero return code with empty output so detection can
    degrade gracefully in non-Windows environments.
    """

    command_list = [str(part) for part in arguments]
    if not command_list:
        return 1, ""

    event = "detect_" + command_list[0].lower().replace("/", "_").replace("\\", "_").replace(
        ".", "_"
    )
    result = exec_utils.run_command(command_list, event=event)

    output = result.stdout or result.stderr or ""
    return result.returncode, output


def main() -> int:
    """!
    @brief Module entry point returning inventory as JSON to stdout.
    """

    payload = gather_office_inventory()
    try:
        sys.stdout.write(json.dumps(payload, default=str))
    except Exception:
        sys.stdout.write("{}")
        return 1
    return 0


if __name__ == "__main__":  # pragma: no cover - manual execution
    raise SystemExit(main())


def gather_running_office_processes() -> list[dict[str, str]]:
    """!
    @brief Inspect running processes for Office executables via ``tasklist``.
    @details Output is filtered to the executables referenced in the
    specification so downstream planners can prompt for graceful shutdowns
    before uninstall operations commence.
    """

    code, output = _run_command(["tasklist", "/FO", "CSV"])
    if code != 0 and not output:
        return []

    processes: list[dict[str, str]] = []
    reader = csv.reader(line.strip("\ufeff") for line in output.splitlines() if line.strip())

    for row in reader:
        if not row:
            continue
        header = row[0].strip().lower()
        if header == "image name":
            continue
        name = row[0].strip()
        if name.lower() not in _OFFICE_PROCESS_TARGETS:
            continue
        entry: dict[str, str] = {"name": name}
        if len(row) > 1:
            entry["pid"] = row[1].strip()
        if len(row) > 2:
            entry["session"] = row[2].strip()
        if len(row) > 3:
            entry["session_id"] = row[3].strip()
        if len(row) > 4:
            entry["memory"] = row[4].strip()
        processes.append(entry)

    return processes


def gather_office_services() -> list[dict[str, str]]:
    """!
    @brief Enumerate Office-related Windows services via ``sc query``.
    @details The collected state helps diagnose Click-to-Run agent activity and
    licensing daemons prior to remediation.
    """

    code, output = _run_command(["sc", "query", "state=", "all"])
    if code != 0 and not output:
        return []

    services: list[dict[str, str]] = []
    current_name: str | None = None

    for raw_line in output.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if line.upper().startswith("SERVICE_NAME"):
            _, _, value = line.partition(":")
            candidate = value.strip()
            current_name = candidate if candidate else None
            continue
        if line.upper().startswith("STATE") and current_name:
            _, _, value = line.partition(":")
            state_detail = value.strip()
            state_parts = state_detail.split()
            state = (
                state_parts[1] if len(state_parts) > 1 else state_parts[0] if state_parts else ""
            )
            if current_name.lower() in _SERVICE_TARGETS:
                services.append(
                    {
                        "name": current_name,
                        "state": state,
                        "details": state_detail,
                    }
                )
            current_name = None

    return services


def gather_office_tasks() -> list[dict[str, str]]:
    """!
    @brief Query scheduled tasks associated with Office maintenance.
    @details Uses ``schtasks`` to surface telemetry, licensing, and background
    handlers that may interfere with uninstall flows.
    """

    code, output = _run_command(["schtasks", "/Query", "/FO", "CSV"])
    if code != 0 and not output:
        return []

    tasks: list[dict[str, str]] = []
    reader = csv.reader(line.strip("\ufeff") for line in output.splitlines() if line.strip())

    for row in reader:
        if not row:
            continue
        task_name = row[0].strip()
        if task_name.lower() == "taskname":
            continue
        if not any(task_name.startswith(prefix) for prefix in _TASK_PREFIXES):
            continue
        entry: dict[str, str] = {"task": task_name}
        if len(row) > 1:
            entry["next_run_time"] = row[1].strip()
        if len(row) > 2:
            entry["status"] = row[2].strip()
        entry["known"] = task_name in _KNOWN_TASK_NAMES
        tasks.append(entry)

    return tasks


def gather_activation_state() -> dict[str, Any]:
    """!
    @brief Inspect activation metadata from the Office Software Protection Platform.
    @details Converts registry values to JSON-friendly primitives for inclusion
    in diagnostics and archival logs.
    """

    registry_path = constants.OSPP_REGISTRY_PATH
    hive_name, _, relative_path = registry_path.partition("\\")
    if not hive_name or not relative_path:
        return {}

    hive = constants.REGISTRY_ROOTS.get(hive_name.upper())
    if hive is None:
        return {}

    values = _read_values_with_fallback(hive, relative_path)
    if not values:
        return {}

    serialised: dict[str, Any] = {}
    for key, value in values.items():
        if isinstance(value, (str, int, float, bool)) or value is None:
            serialised[str(key)] = value
        else:
            serialised[str(key)] = str(value)

    return {"path": registry_path, "values": serialised}


def gather_registry_residue() -> list[dict[str, str]]:
    """!
    @brief Identify registry hives that likely require cleanup.
    @details The returned list mirrors OffScrub residue heuristics so planners
    can schedule deletions alongside filesystem cleanup once uninstalls complete.
    """

    entries: list[dict[str, str]] = []

    for hive, path in constants.REGISTRY_RESIDUE_PATHS:
        if not _key_exists_with_fallback(hive, path):
            continue
        entries.append({"path": _compose_handle(hive, path)})

    return entries
