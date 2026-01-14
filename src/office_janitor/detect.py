"""!
@brief Detection helpers for MSI and Click-to-Run Office deployments.
@details Reads structured metadata from :mod:`office_janitor.constants`, probes
registry hives, and returns structured :class:`DetectedInstallation` records that
contain uninstall handles, source type, and channel information.
"""

from __future__ import annotations

import concurrent.futures
import csv
import json
import logging
import os
import shlex
import sys
import threading
import time
from collections.abc import Callable, Iterable, Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any, TypeVar

from . import constants, elevation, exec_utils, logging_ext, registry_tools, spinner

_LOGGER = logging.getLogger(__name__)

# Progress update intervals for long-running operations (in seconds)
_PROGRESS_INTERVALS = (30, 60, 120, 300, 600, 1800)  # 30s, 1m, 2m, 5m, 10m, 30m

_T = TypeVar("_T")


def _wait_with_progress(
    future: concurrent.futures.Future[_T],
    task_name: str,
    report_fn: Callable[[str], None] | None = None,
    poll_interval: float = 0.5,
    expected_time: float | None = None,
    use_spinner: bool = True,
) -> _T:
    """!
    @brief Wait on a future while showing progress via the global spinner.
    @details Uses the global spinner module to display the current task.
    Periodically emits log messages at predefined intervals (30s, 1m, 2m,
    5m, 10m, 30m) to reassure users that long-running operations have not stalled.

    @param future The concurrent.futures.Future to wait on.
    @param task_name Descriptive name for logging (e.g., "WMI probes").
    @param report_fn Optional callback for status messages; falls back to _LOGGER.info.
    @param poll_interval How often (in seconds) to check the future's status.
    @param expected_time Optional expected duration (for documentation; spinner handles display).
    @param use_spinner Whether to update the global spinner with this task.
    @return The result of the future once complete.
    @raises Any exception raised by the future's underlying task.
    @raises KeyboardInterrupt if cancellation is requested.
    """
    start_time = time.monotonic()
    next_intervals = list(_PROGRESS_INTERVALS)  # Mutable copy to track which have fired

    def _emit_log(msg: str) -> None:
        """Emit a log message."""
        if report_fn is not None:
            report_fn(msg)
        else:
            _LOGGER.info(msg)

    # Set the global spinner task
    if use_spinner:
        spinner.set_task(task_name)

    try:
        while not future.done():
            # Check for cancellation (allows Ctrl+C to work)
            spinner.check_cancelled()

            try:
                future.result(timeout=poll_interval)
                break  # Completed successfully
            except concurrent.futures.TimeoutError:
                elapsed = time.monotonic() - start_time

                # Check if we've crossed any progress thresholds for log messages
                while next_intervals and elapsed >= next_intervals[0]:
                    threshold = next_intervals.pop(0)
                    elapsed_str = spinner._format_elapsed(threshold)
                    _emit_log(f"Still working on {task_name}... ({elapsed_str} elapsed)")

    finally:
        # Clear spinner task when done (but don't stop the spinner thread)
        if use_spinner:
            spinner.clear_task()

    return future.result()


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
    @details WARNING: This is slow (30-120+ seconds) because wmic enumerates all products.
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
        # Exclude non-Office products that match keywords (e.g., Aspire.ProjectTemplates)
        name_lower = name.lower()
        if any(excl in name_lower for excl in ("aspire", "template", "sdk", "visual studio")):
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
        # Exclude non-Office products that match keywords (e.g., Aspire.ProjectTemplates)
        name_lower = name.lower()
        if any(excl in name_lower for excl in ("aspire", "template", "sdk", "visual studio")):
            continue
        results[product_code.upper()] = {
            "product": name or product_code,
            "version": version,
            "install_location": install_path,
            "probe": "powershell",
        }

    return results


def detect_appx_packages() -> list[dict[str, object]]:
    """!
    @brief Detect installed Office AppX/MSIX packages (modern Windows apps).
    @details Uses PowerShell Get-AppxPackage to enumerate installed modern
    Office apps. These are separate from MSI and Click-to-Run installations.
    """

    # PowerShell script to get Office-related AppX packages
    script = (
        "$ErrorActionPreference='SilentlyContinue';"
        "$packages = Get-AppxPackage -AllUsers -ErrorAction SilentlyContinue | "
        "Where-Object { "
        "$_.Name -like '*Microsoft.Office*' -or "
        "$_.Name -like '*Microsoft.365*' -or "
        "$_.Name -like '*Microsoft.MicrosoftOffice*' -or "
        "$_.Name -like '*Microsoft.Outlook*' -or "
        "$_.Name -like '*Microsoft.Excel*' -or "
        "$_.Name -like '*Microsoft.Word*' -or "
        "$_.Name -like '*Microsoft.PowerPoint*' -or "
        "$_.Name -like '*Microsoft.OneNote*' -or "
        "$_.Name -like '*Microsoft.Visio*' -or "
        "$_.Name -like '*Microsoft.Project*' -or "
        "$_.Name -like '*Microsoft.Access*' "
        "} | "
        "Select-Object Name,PackageFullName,Version,Architecture,InstallLocation,Publisher;"
        "if($packages){$packages|ConvertTo-Json -Compress}else{''}"
    )

    code, output = _run_command(["powershell", "-NoProfile", "-Command", script])
    if code != 0 or not output.strip():
        return []

    text = output.strip()
    try:
        payload = json.loads(text)
    except json.JSONDecodeError:
        return []

    records = payload if isinstance(payload, list) else [payload]
    results: list[dict[str, object]] = []

    # Exclusion patterns for non-Office apps
    exclusions = (
        "teams",
        "visualstudio",
        "vscode",
        "azure",
        "sql",
        "powershell",
    )

    for record in records:
        if not isinstance(record, Mapping):
            continue

        name = str(record.get("Name") or "").strip()
        if not name:
            continue

        # Skip excluded packages
        name_lower = name.lower()
        if any(excl in name_lower for excl in exclusions):
            continue

        package_full_name = str(record.get("PackageFullName") or "").strip()
        version = str(record.get("Version") or "").strip()
        architecture = str(record.get("Architecture") or "").strip()
        install_location = str(record.get("InstallLocation") or "").strip()
        publisher = str(record.get("Publisher") or "").strip()

        results.append(
            {
                "name": name,
                "package_full_name": package_full_name,
                "version": version,
                "architecture": architecture,
                "install_location": install_location,
                "publisher": publisher,
                "source": "AppX",
            }
        )

    return results


def detect_uninstall_entries() -> list[dict[str, object]]:
    """!
    @brief Detect Office entries from Windows Control Panel uninstall registry.
    @details Scans the standard Uninstall registry keys that populate the
    "Programs and Features" / "Apps & Features" control panel. This catches
    Office installations that may not be in the known MSI product map.
    """

    results: list[dict[str, object]] = []
    seen_handles: set[str] = set()

    # Use the registry_tools function to find Office-like uninstall entries
    for hive, key_path, values in registry_tools.iter_office_uninstall_entries(
        constants.MSI_UNINSTALL_ROOTS
    ):
        handle = f"HKLM\\{key_path}" if hive == constants.HKLM else f"HKCU\\{key_path}"
        if handle in seen_handles:
            continue
        seen_handles.add(handle)

        display_name = str(values.get("DisplayName") or "").strip()
        if not display_name:
            continue

        # Extract useful fields
        display_version = str(values.get("DisplayVersion") or "").strip()
        publisher = str(values.get("Publisher") or "").strip()
        install_location = str(values.get("InstallLocation") or "").strip()
        uninstall_string = str(values.get("UninstallString") or "").strip()
        quiet_uninstall = str(values.get("QuietUninstallString") or "").strip()
        install_date = str(values.get("InstallDate") or "").strip()

        # Try to extract product code from key path
        product_code = key_path.rsplit("\\", 1)[-1] if "\\" in key_path else key_path

        results.append(
            {
                "display_name": display_name,
                "version": display_version,
                "publisher": publisher,
                "install_location": install_location,
                "uninstall_string": uninstall_string,
                "quiet_uninstall_string": quiet_uninstall,
                "install_date": install_date,
                "product_code": product_code,
                "registry_handle": handle,
                "source": "ControlPanel",
            }
        )

    return results


def detect_msi_installations(
    *,
    skip_slow_probes: bool = False,
    precomputed_fallbacks: dict[str, dict[str, Any]] | None = None,
) -> list[DetectedInstallation]:
    """!
    @brief Inspect the registry and return metadata for MSI-based Office installs.
    @param skip_slow_probes If True, skip WMI/PowerShell probes that can take 60-120+ seconds.
    @param precomputed_fallbacks Pre-collected WMI/PS probe results (avoids re-running probes).
    """

    installations: list[DetectedInstallation] = []
    seen_handles: set[str] = set()
    seen_codes: set[str] = set()

    fallback_sources: dict[str, dict[str, Any]] = {}

    if precomputed_fallbacks is not None:
        # Use pre-computed fallback data (probes ran externally in parallel)
        fallback_sources.update(precomputed_fallbacks)
    elif skip_slow_probes:
        # Skip the extremely slow WMI queries
        _LOGGER.debug("Skipping slow MSI probes (WMI/PowerShell)")
    else:
        # Run both slow probes in parallel to reduce total wait time
        with concurrent.futures.ThreadPoolExecutor(
            max_workers=2, thread_name_prefix="msi_probe"
        ) as executor:
            wmi_future = executor.submit(_probe_msi_wmi)
            ps_future = executor.submit(_probe_msi_powershell)
            # Merge results as they complete
            _merge_fallback_metadata(fallback_sources, wmi_future.result())
            _merge_fallback_metadata(fallback_sources, ps_future.result())

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
        properties2: dict[str, object] = {
            "display_name": display_name,
            "display_version": version,
        }
        if install_location:
            properties2["install_location"] = install_location
        probe = metadata.get("probe")
        if probe:
            properties2["supplemental_probes"] = [str(probe)]

        installations.append(
            DetectedInstallation(
                source="MSI",
                product=display_name,
                version=version or "unknown",
                architecture=architecture,
                uninstall_handles=handles,
                channel="MSI",
                product_code=product_code,
                properties=properties2,
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


def gather_office_inventory(
    *,
    limited_user: bool | None = None,
    progress_callback: Callable[[str, str], None] | None = None,
    parallel: bool = True,
    fast_mode: bool = False,
) -> dict[str, object]:
    """!
    @brief Aggregate MSI, C2R, and ancillary signals into an inventory payload.
    @param limited_user If True, attempt de-elevated probes for user context.
    @param progress_callback Optional callback(phase, status) for progress reporting.
           phase is a description, status is "start", "ok", "skip", or "fail".
    @param parallel If True, run independent detection tasks in parallel threads.
    @param fast_mode If True, skip slow WMI/PowerShell probes (reduces MSI detection from
           60-120+ seconds to under 1 second, but may miss some edge-case installations).
    """

    # Start the spinner thread (for use during slow operations only)
    spinner.start_spinner_thread()

    # Thread-safe progress reporting with spinner updates
    _report_lock = threading.Lock()

    def _report(phase: str, status: str = "start") -> None:
        """Update spinner task and optionally call progress callback."""
        with _report_lock:
            # Update spinner with current phase (only on start, not completion)
            if status == "start":
                spinner.set_task(phase)
            if progress_callback:
                progress_callback(phase, status)

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
        _report("De-elevating for user context probe")
        result = elevation.run_as_limited_user(
            [sys.executable, "-m", "office_janitor.detect"],
            event="detect_deelevate",
            env_overrides={"OFFICE_JANITOR_DEELEVATED": "1"},
        )
        if result.returncode == 0 and result.stdout:
            try:
                parsed = json.loads(result.stdout)
                if isinstance(parsed, dict):
                    _report("De-elevated probe", "ok")
                    return parsed
            except json.JSONDecodeError:
                _report("De-elevated probe", "fail")
                human_logger.warning(
                    "Failed to parse limited-user detection output; falling back to current "
                    "context."
                )

    # Context is quick and needed first
    _report("Checking execution context")
    context_info = {
        "user": elevation.current_username(),
        "is_admin": elevation.is_admin(),
    }
    _report("Checking execution context", "ok")

    # Start slow WMI/PS probes immediately (they run in background while other tasks execute)
    wmi_future: concurrent.futures.Future[dict[str, dict[str, Any]]] | None = None
    ps_future: concurrent.futures.Future[dict[str, dict[str, Any]]] | None = None
    probe_executor: concurrent.futures.ThreadPoolExecutor | None = None

    if not fast_mode and parallel:
        _report("Starting WMI/PowerShell probes (background)")
        probe_executor = concurrent.futures.ThreadPoolExecutor(
            max_workers=2, thread_name_prefix="msi_probe"
        )
        wmi_future = probe_executor.submit(_probe_msi_wmi)
        ps_future = probe_executor.submit(_probe_msi_powershell)

    # Define detection tasks for parallel execution
    def _detect_msi(
        precomputed_fallbacks: dict[str, dict[str, Any]] | None = None,
    ) -> list[dict[str, object]]:
        if fast_mode:
            _report("Scanning MSI-based installations (fast mode)")
            result = [entry.to_dict() for entry in detect_msi_installations(skip_slow_probes=True)]
            _report("Scanning MSI-based installations (fast mode)", "ok")
        else:
            _report("Scanning MSI-based installations")
            result = [
                entry.to_dict()
                for entry in detect_msi_installations(
                    skip_slow_probes=True,  # We handle probes externally
                    precomputed_fallbacks=precomputed_fallbacks,
                )
            ]
            _report("Scanning MSI-based installations", "ok")
        return result

    def _detect_c2r() -> list[dict[str, object]]:
        _report("Scanning Click-to-Run installations")
        result = [entry.to_dict() for entry in detect_c2r_installations()]
        _report("Scanning Click-to-Run installations", "ok")
        return result

    def _detect_processes() -> list[dict[str, object]]:
        _report("Enumerating running Office processes")
        raw = gather_running_office_processes()
        result: list[dict[str, object]] = [dict(r) for r in raw]
        _report("Enumerating running Office processes", "ok")
        return result

    def _detect_services() -> list[dict[str, object]]:
        _report("Checking Office services")
        raw = gather_office_services()
        result: list[dict[str, object]] = [dict(r) for r in raw]
        _report("Checking Office services", "ok")
        return result

    def _detect_tasks() -> list[dict[str, object]]:
        _report("Checking scheduled tasks")
        raw = gather_office_tasks()
        result: list[dict[str, object]] = [dict(r) for r in raw]
        _report("Checking scheduled tasks", "ok")
        return result

    def _detect_appx() -> list[dict[str, object]]:
        _report("Scanning AppX/MSIX packages")
        raw = detect_appx_packages()
        result: list[dict[str, object]] = [dict(r) for r in raw]
        _report("Scanning AppX/MSIX packages", "ok")
        return result

    def _detect_uninstall_entries() -> list[dict[str, object]]:
        _report("Scanning Control Panel uninstall entries")
        raw = detect_uninstall_entries()
        result: list[dict[str, object]] = [dict(r) for r in raw]
        _report("Scanning Control Panel uninstall entries", "ok")
        return result

    def _detect_activation() -> dict[str, object]:
        _report("Gathering activation/licensing state")
        result = gather_activation_state()
        _report("Gathering activation/licensing state", "ok")
        return result

    def _detect_registry() -> list[dict[str, object]]:
        _report("Scanning registry for residue")
        raw = gather_registry_residue()
        result: list[dict[str, object]] = [dict(r) for r in raw]
        _report("Scanning registry for residue", "ok")
        return result

    def _detect_filesystem() -> list[dict[str, object]]:
        _report("Scanning filesystem paths")
        fs_entries: list[dict[str, object]] = []
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
            fs_entries.append(
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
            fs_entries.append(entry)
            seen_paths.add(path_str)

        _report("Scanning filesystem paths", "ok")
        return fs_entries

    # Run detection tasks in parallel or sequentially
    if parallel:
        try:
            with concurrent.futures.ThreadPoolExecutor(
                max_workers=14, thread_name_prefix="detect"
            ) as executor:
                # Start all fast tasks immediately (WMI/PS probes already running in background)
                c2r_future = executor.submit(_detect_c2r)
                processes_future = executor.submit(_detect_processes)
                services_future = executor.submit(_detect_services)
                tasks_future = executor.submit(_detect_tasks)
                appx_future = executor.submit(_detect_appx)
                uninstall_future = executor.submit(_detect_uninstall_entries)
                activation_future = executor.submit(_detect_activation)
                registry_future = executor.submit(_detect_registry)
                filesystem_future = executor.submit(_detect_filesystem)

                # Use interruptible waiting for all futures
                futures = {
                    "c2r": c2r_future,
                    "processes": processes_future,
                    "services": services_future,
                    "tasks": tasks_future,
                    "appx": appx_future,
                    "uninstall": uninstall_future,
                    "activation": activation_future,
                    "registry": registry_future,
                    "filesystem": filesystem_future,
                }

                results = spinner.wait_for_futures(futures, poll_interval=0.1)

                c2r_list = results["c2r"]
                processes_list = results["processes"]
                services_list = results["services"]
                tasks_list = results["tasks"]
                appx_list = results["appx"]
                uninstall_list = results["uninstall"]
                activation_info = results["activation"]
                registry_residue = results["registry"]
                filesystem_list = results["filesystem"]

                # Now that all fast tasks are done and console is quiet,
                # wait for WMI/PS probes with spinner (safe to use spinner now)
                probe_fallbacks: dict[str, dict[str, Any]] = {}
                if wmi_future is not None and ps_future is not None:
                    try:
                        _report("Waiting for WMI/PowerShell probes")
                        human_logger.info(
                            "WMI/PowerShell probes may take 1-3 minutes depending on system..."
                        )
                        # Use progress-aware waiting with spinner and time estimate
                        # WMI typically takes 60-120s, PS takes 30-60s
                        wmi_result = _wait_with_progress(
                            wmi_future,
                            "WMI probes",
                            lambda msg: human_logger.info(msg),
                            expected_time=90.0,  # ~1.5 min typical
                            use_spinner=True,
                        )
                        _merge_fallback_metadata(probe_fallbacks, wmi_result)
                        ps_result = _wait_with_progress(
                            ps_future,
                            "PowerShell probes",
                            lambda msg: human_logger.info(msg),
                            expected_time=45.0,  # ~45s typical
                            use_spinner=True,
                        )
                        _merge_fallback_metadata(probe_fallbacks, ps_result)
                        _report("Waiting for WMI/PowerShell probes", "ok")
                    except (KeyboardInterrupt, concurrent.futures.CancelledError):
                        _report("Waiting for WMI/PowerShell probes", "skip")
                        raise

                # Now run MSI detection with probe results
                msi_list = _detect_msi(probe_fallbacks if probe_fallbacks else None)

        except KeyboardInterrupt:
            _LOGGER.info("Detection interrupted by user")
            # Cancel any pending futures
            if probe_executor is not None:
                probe_executor.shutdown(wait=False, cancel_futures=True)
            raise

        # Clean up probe executor
        if probe_executor is not None:
            probe_executor.shutdown(wait=False)
    else:
        # Sequential fallback (with probes if not fast_mode)
        probe_fallbacks_seq: dict[str, dict[str, Any]] = {}
        if not fast_mode:
            _merge_fallback_metadata(probe_fallbacks_seq, _probe_msi_wmi())
            _merge_fallback_metadata(probe_fallbacks_seq, _probe_msi_powershell())
        msi_list = _detect_msi(probe_fallbacks_seq if probe_fallbacks_seq else None)
        c2r_list = _detect_c2r()
        processes_list = _detect_processes()
        services_list = _detect_services()
        tasks_list = _detect_tasks()
        appx_list = _detect_appx()
        uninstall_list = _detect_uninstall_entries()
        activation_info = _detect_activation()
        registry_residue = _detect_registry()
        filesystem_list = _detect_filesystem()

    inventory: dict[str, object] = {
        "context": context_info,
        "msi": msi_list,
        "c2r": c2r_list,
        "appx": appx_list,
        "uninstall_entries": uninstall_list,
        "filesystem": filesystem_list,
        "processes": processes_list,
        "services": services_list,
        "tasks": tasks_list,
        "activation": activation_info,
        "registry": registry_residue,
    }

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


# ---------------------------------------------------------------------------
# Temporary ARP Entry Management (for orphaned MSI products)
# ---------------------------------------------------------------------------

TEMP_ARP_KEY_PREFIX = "OFFICE_TEMP"
"""!
@brief Prefix for temporary ARP entries created by the scrubber.
@details These entries enable msiexec uninstall for products that have lost
their Add/Remove Programs registration but still have WI metadata.
"""


def _generate_temp_arp_key(product_code: str, version: str = "16") -> str:
    """!
    @brief Generate a unique temporary ARP key name.
    @param product_code The MSI product code.
    @param version Office version (15, 16, etc.).
    @returns Key name like "OFFICE_TEMP.{GUID}".
    """
    # Normalize the product code
    code = product_code.strip().upper()
    if not code.startswith("{"):
        code = f"{{{code}}}"
    if not code.endswith("}"):
        code = f"{code}}}"
    return f"{TEMP_ARP_KEY_PREFIX}{version}.{code}"


def create_temp_arp_entry(
    product_code: str,
    product_name: str,
    *,
    version: str = "16",
    install_location: str | None = None,
    dry_run: bool = False,
) -> str | None:
    """!
    @brief Create a temporary ARP entry for an orphaned MSI product.
    @details VBS equivalent: arrTmpSKUs population in FindInstalledOProducts.
        This allows msiexec to find and uninstall products that have lost
        their Add/Remove Programs registration but still exist in WI metadata.
    @param product_code The MSI product code (GUID).
    @param product_name Display name for the product.
    @param version Office major version (15, 16, etc.).
    @param install_location Optional install path.
    @param dry_run If True, only log what would be created.
    @returns The registry path of the created entry, or None on failure.
    """
    human_logger = logging_ext.get_human_logger()

    # Validate product code format
    code = product_code.strip().upper()
    if not code.startswith("{"):
        code = f"{{{code}}}"

    # Generate key name
    key_name = _generate_temp_arp_key(code, version)

    # Build registry path
    base_path = r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
    full_path = f"HKLM\\{base_path}\\{key_name}"

    if dry_run:
        human_logger.info("[DRY-RUN] Would create temp ARP entry: %s", full_path)
        return full_path

    human_logger.info("Creating temporary ARP entry: %s", key_name)

    # Create the registry key with required values
    try:
        import winreg

        with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, f"{base_path}\\{key_name}") as key:
            winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, product_name)
            winreg.SetValueEx(
                key,
                "UninstallString",
                0,
                winreg.REG_SZ,
                f"msiexec.exe /X{code} /qn",
            )
            winreg.SetValueEx(key, "SystemComponent", 0, winreg.REG_DWORD, 0)
            winreg.SetValueEx(key, "Publisher", 0, winreg.REG_SZ, "Microsoft Corporation")
            winreg.SetValueEx(
                key, "Comments", 0, winreg.REG_SZ, "Temporary entry for Office cleanup"
            )
            # Mark as temporary so we can clean it up later
            winreg.SetValueEx(key, "OfficeJanitorTemp", 0, winreg.REG_DWORD, 1)
            if install_location:
                winreg.SetValueEx(key, "InstallLocation", 0, winreg.REG_SZ, install_location)

        human_logger.debug("Created temp ARP entry: %s", full_path)
        return full_path

    except OSError as exc:
        human_logger.warning("Failed to create temp ARP entry %s: %s", key_name, exc)
        return None


def cleanup_temp_arp_entries(*, dry_run: bool = False) -> int:
    """!
    @brief Remove all temporary ARP entries created by the scrubber.
    @details Cleans up entries with the OFFICE_TEMP prefix or OfficeJanitorTemp marker.
    @param dry_run If True, only log what would be deleted.
    @returns Number of entries removed.
    """
    human_logger = logging_ext.get_human_logger()
    removed = 0

    try:
        import winreg

        base_path = r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"

        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, base_path, 0, winreg.KEY_READ) as base_key:
            index = 0
            keys_to_delete: list[str] = []

            while True:
                try:
                    subkey_name = winreg.EnumKey(base_key, index)
                    index += 1

                    # Check if it's a temp key by prefix
                    if subkey_name.startswith(TEMP_ARP_KEY_PREFIX):
                        keys_to_delete.append(subkey_name)
                        continue

                    # Also check for our marker value
                    try:
                        with winreg.OpenKey(base_key, subkey_name, 0, winreg.KEY_READ) as subkey:
                            try:
                                value, _ = winreg.QueryValueEx(subkey, "OfficeJanitorTemp")
                                if value:
                                    keys_to_delete.append(subkey_name)
                            except FileNotFoundError:
                                pass
                    except OSError:
                        pass

                except OSError:
                    break

        # Delete the identified keys
        for key_name in keys_to_delete:
            full_path = f"HKLM\\{base_path}\\{key_name}"
            if dry_run:
                human_logger.info("[DRY-RUN] Would delete temp ARP: %s", key_name)
                removed += 1
            else:
                try:
                    winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, f"{base_path}\\{key_name}")
                    human_logger.debug("Deleted temp ARP entry: %s", key_name)
                    removed += 1
                except OSError as exc:
                    human_logger.warning("Failed to delete temp ARP %s: %s", key_name, exc)

    except OSError as exc:
        human_logger.debug("Failed to enumerate ARP entries: %s", exc)

    if removed:
        human_logger.info("Cleaned up %d temporary ARP entries", removed)

    return removed


def find_orphaned_wi_products() -> list[dict[str, str]]:
    """!
    @brief Find MSI products in WI metadata without ARP entries.
    @details VBS equivalent: fTryReconcile logic in FindInstalledOProducts.
        Scans Windows Installer product registry for Office products that
        have no corresponding Add/Remove Programs entry.
    @returns List of orphaned products with product_code, name, and version.
    """
    from . import guid_utils

    human_logger = logging_ext.get_human_logger()
    orphans: list[dict[str, str]] = []

    try:
        import winreg

        # Get all ARP entries for comparison
        arp_codes: set[str] = set()
        arp_base = r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, arp_base, 0, winreg.KEY_READ) as arp_key:
                index = 0
                while True:
                    try:
                        subkey = winreg.EnumKey(arp_key, index)
                        index += 1
                        # Check if subkey looks like a product code
                        if subkey.startswith("{") and subkey.endswith("}"):
                            arp_codes.add(subkey.upper())
                    except OSError:
                        break
        except OSError:
            pass

        # Scan WI Products
        wi_products_path = r"SOFTWARE\\Classes\\Installer\\Products"
        try:
            with winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE, wi_products_path, 0, winreg.KEY_READ
            ) as wi_key:
                index = 0
                while True:
                    try:
                        compressed = winreg.EnumKey(wi_key, index)
                        index += 1

                        # Try to expand the compressed GUID
                        try:
                            product_code = guid_utils.expand_guid(compressed)
                        except guid_utils.GuidError:
                            continue

                        # Check if this is an Office product
                        if not guid_utils.is_office_guid(product_code):
                            continue

                        # Check if it has an ARP entry
                        if product_code.upper() in arp_codes:
                            continue

                        # Read product name from WI
                        product_name = "Unknown Office Product"
                        try:
                            with winreg.OpenKey(wi_key, compressed, 0, winreg.KEY_READ) as prod_key:
                                try:
                                    product_name, _ = winreg.QueryValueEx(prod_key, "ProductName")
                                except FileNotFoundError:
                                    pass
                        except OSError:
                            pass

                        orphans.append(
                            {
                                "product_code": product_code,
                                "compressed_guid": compressed,
                                "name": product_name,
                                "version": guid_utils.get_office_version_from_guid(product_code)
                                or "unknown",
                            }
                        )

                    except OSError:
                        break

        except OSError:
            pass

    except ImportError:
        human_logger.debug("winreg not available, skipping orphan scan")

    if orphans:
        human_logger.info("Found %d orphaned WI products", len(orphans))

    return orphans


def create_arp_entries_for_orphans(*, dry_run: bool = False) -> int:
    """!
    @brief Create temporary ARP entries for all orphaned WI products.
    @details Enables msiexec uninstall for products that lost their ARP registration.
    @param dry_run If True, only log what would be created.
    @returns Number of entries created.
    """
    orphans = find_orphaned_wi_products()
    created = 0

    for orphan in orphans:
        result = create_temp_arp_entry(
            orphan["product_code"],
            orphan["name"],
            version=orphan.get("version", "16"),
            dry_run=dry_run,
        )
        if result:
            created += 1

    return created


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
