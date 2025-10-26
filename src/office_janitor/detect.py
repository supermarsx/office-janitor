"""!
@brief Detection helpers for MSI and Click-to-Run Office deployments.
@details Reads structured metadata from :mod:`office_janitor.constants`, probes
registry hives, and returns structured :class:`DetectedInstallation` records that
contain uninstall handles, source type, and channel information.
"""
from __future__ import annotations

import csv
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Tuple

from . import constants, registry_tools


_OFFICE_PROCESS_TARGETS = tuple(
    sorted(
        {name.lower() for name in constants.DEFAULT_OFFICE_PROCESSES}
        | {"mspub.exe", "teams.exe"}
    )
)
"""!
@brief Known Office executables monitored during detection.
"""

_SERVICE_TARGETS = tuple(
    sorted({name.lower() for name in constants.KNOWN_SERVICES} | {"osppsvc"})
)
"""!
@brief Services associated with Office provisioning and licensing.
"""

_TASK_PREFIXES = (r"\\Microsoft\\Office\\", r"\\Microsoft\\OfficeSoftwareProtectionPlatform\\")
"""!
@brief Scheduled task prefixes that indicate Office automation jobs.
"""

_KNOWN_TASK_NAMES = {
    task if task.startswith("\\") else f"\\{task}"
    for task in constants.KNOWN_SCHEDULED_TASKS
}
"""!
@brief Explicit scheduled task identifiers from the specification.
"""

_REGISTRY_RESIDUE_TEMPLATES: Tuple[Tuple[int, str], ...] = (
    (constants.HKLM, r"SOFTWARE\\Microsoft\\Office"),
    (constants.HKLM, r"SOFTWARE\\WOW6432Node\\Microsoft\\Office"),
    (constants.HKCU, r"SOFTWARE\\Microsoft\\Office"),
    (constants.HKLM, r"SOFTWARE\\Microsoft\\ClickToRun"),
    (constants.HKLM, r"SOFTWARE\\Microsoft\\Office\\Common"),
    (constants.HKLM, r"SOFTWARE\\Microsoft\\OfficeSoftwareProtectionPlatform"),
)
"""!
@brief Registry hives monitored for residue cleanup opportunities.
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
    uninstall_handles: Tuple[str, ...]
    channel: str
    product_code: str | None = None
    release_ids: Tuple[str, ...] = ()
    properties: Mapping[str, object] | None = None

    def to_dict(self) -> Dict[str, object]:
        """!
        @brief Convert the dataclass to a JSON-serialisable dictionary.
        """

        payload: Dict[str, object] = {
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
        return payload


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


def detect_msi_installations() -> List[DetectedInstallation]:
    """!
    @brief Inspect the registry and return metadata for MSI-based Office installs.
    """

    installations: List[DetectedInstallation] = []
    seen_handles: set[str] = set()

    for product_code, metadata in constants.MSI_PRODUCT_MAP.items():
        registry_roots: Iterable[Tuple[int, str]] = metadata.get(
            "registry_roots", constants.MSI_UNINSTALL_ROOTS
        )
        for hive, base_key in registry_roots:
            key_path = f"{base_key}\\{product_code}"
            values = registry_tools.read_values(hive, key_path)
            if not values:
                continue

            handle = _compose_handle(hive, key_path)
            if handle in seen_handles:
                continue

            display_name = str(values.get("DisplayName") or metadata.get("product") or product_code)
            display_version = str(values.get("DisplayVersion") or "")
            uninstall_string = str(values.get("UninstallString") or "")
            family = constants.resolve_msi_family(product_code) or str(metadata.get("family", ""))

            properties: Dict[str, object] = {
                "display_name": display_name,
                "display_version": display_version,
            }
            if uninstall_string:
                properties["uninstall_string"] = uninstall_string
            properties["supported_versions"] = list(metadata.get("supported_versions", ()))
            properties["edition"] = metadata.get("edition", "")
            if family:
                properties["family"] = family

            installations.append(
                DetectedInstallation(
                    source="MSI",
                    product=str(metadata.get("product", display_name)),
                    version=str(metadata.get("version", "unknown")),
                    architecture=str(metadata.get("architecture", "unknown")),
                    uninstall_handles=(handle,),
                    channel="MSI",
                    product_code=product_code,
                    properties=properties,
                )
            )
            seen_handles.add(handle)

    return installations


def detect_c2r_installations() -> List[DetectedInstallation]:
    """!
    @brief Probe Click-to-Run configuration to describe installed suites.
    """

    installations: List[DetectedInstallation] = []

    for hive, config_path in constants.C2R_CONFIGURATION_KEYS:
        config_values = registry_tools.read_values(hive, config_path)
        if not config_values:
            continue

        raw_release_ids = str(config_values.get("ProductReleaseIds", "")).split(",")
        release_ids = tuple(sorted(rid.strip() for rid in raw_release_ids if rid.strip()))
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
        channel = _friendly_channel(str(channel_identifier) if channel_identifier else None)
        package_guid = str(config_values.get("PackageGUID") or "")
        install_path = str(config_values.get("InstallPath") or "")

        for release_id in release_ids:
            product_metadata = constants.C2R_PRODUCT_RELEASES.get(release_id)
            product_name = str(product_metadata.get("product", release_id)) if product_metadata else release_id
            supported_versions = tuple(
                str(v) for v in (product_metadata or {}).get("supported_versions", ())
            )
            supported_architectures = tuple(
                str(a) for a in (product_metadata or {}).get("architectures", ())
            )
            family = constants.resolve_c2r_family(release_id) or str(
                (product_metadata or {}).get("family", "")
            )
            uninstall_handles = [_compose_handle(hive, config_path)]

            registry_paths = (product_metadata or {}).get("registry_paths", {})
            release_roots: Iterable[Tuple[int, str]] = registry_paths.get(
                "product_release_ids", constants.C2R_PRODUCT_RELEASE_ROOTS
            )
            for rel_hive, rel_base in release_roots:
                release_key = f"{rel_base}\\{release_id}"
                if registry_tools.key_exists(rel_hive, release_key):
                    uninstall_handles.append(_compose_handle(rel_hive, release_key))

            properties: Dict[str, object] = {
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

            installations.append(
                DetectedInstallation(
                    source="C2R",
                    product=product_name,
                    version=version,
                    architecture=architecture,
                    uninstall_handles=tuple(uninstall_handles),
                    channel=channel,
                    release_ids=(release_id,),
                    properties=properties,
                )
            )

    return installations


def gather_office_inventory() -> Dict[str, object]:
    """!
    @brief Aggregate MSI, C2R, and ancillary signals into an inventory payload.
    """

    inventory: Dict[str, object] = {
        "msi": [entry.to_dict() for entry in detect_msi_installations()],
        "c2r": [entry.to_dict() for entry in detect_c2r_installations()],
        "filesystem": [],
        "processes": gather_running_office_processes(),
        "services": gather_office_services(),
        "tasks": gather_office_tasks(),
        "activation": gather_activation_state(),
        "registry": gather_registry_residue(),
    }

    for template in constants.INSTALL_ROOT_TEMPLATES:
        candidate = Path(template["path"])
        try:
            exists = candidate.exists()
        except OSError:
            exists = False
        if not exists:
            continue
        inventory["filesystem"].append(
            {
                "path": str(candidate),
                "architecture": template.get("architecture", "unknown"),
                "release": template.get("release", ""),
                "label": template.get("label", ""),
            }
        )

    return inventory


def reprobe(options: Mapping[str, object] | None = None) -> Dict[str, object]:
    """!
    @brief Re-run Office detection after a scrub pass to check for leftovers.
    @details The optional ``options`` mapping is accepted for parity with future
    targeted detection strategies, but currently serves only as a hook for
    logging and diagnostics. The returned inventory mirrors
    :func:`gather_office_inventory`.
    """

    _ = options  # Options are presently unused but reserved for parity.
    return gather_office_inventory()


def _run_command(arguments: Iterable[str]) -> Tuple[int, str]:
    """!
    @brief Execute a subprocess returning ``(returncode, text_output)``.
    @details Failures caused by missing binaries or platform limitations are
    normalised to a non-zero return code with empty output so detection can
    degrade gracefully in non-Windows environments.
    """

    try:
        completed = subprocess.run(  # noqa: S603 - intentional command execution
            list(arguments),
            check=False,
            capture_output=True,
            text=True,
        )
    except FileNotFoundError:
        return 127, ""
    except OSError:
        return 1, ""
    output = completed.stdout or completed.stderr or ""
    return completed.returncode, output


def gather_running_office_processes() -> List[Dict[str, str]]:
    """!
    @brief Inspect running processes for Office executables via ``tasklist``.
    @details Output is filtered to the executables referenced in the
    specification so downstream planners can prompt for graceful shutdowns
    before uninstall operations commence.
    """

    code, output = _run_command(["tasklist", "/FO", "CSV"])
    if code != 0 and not output:
        return []

    processes: List[Dict[str, str]] = []
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
        entry: Dict[str, str] = {"name": name}
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


def gather_office_services() -> List[Dict[str, str]]:
    """!
    @brief Enumerate Office-related Windows services via ``sc query``.
    @details The collected state helps diagnose Click-to-Run agent activity and
    licensing daemons prior to remediation.
    """

    code, output = _run_command(["sc", "query", "state=", "all"])
    if code != 0 and not output:
        return []

    services: List[Dict[str, str]] = []
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
            state = state_parts[1] if len(state_parts) > 1 else state_parts[0] if state_parts else ""
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


def gather_office_tasks() -> List[Dict[str, str]]:
    """!
    @brief Query scheduled tasks associated with Office maintenance.
    @details Uses ``schtasks`` to surface telemetry, licensing, and background
    handlers that may interfere with uninstall flows.
    """

    code, output = _run_command(["schtasks", "/Query", "/FO", "CSV"])
    if code != 0 and not output:
        return []

    tasks: List[Dict[str, str]] = []
    reader = csv.reader(line.strip("\ufeff") for line in output.splitlines() if line.strip())

    for row in reader:
        if not row:
            continue
        task_name = row[0].strip()
        if task_name.lower() == "taskname":
            continue
        if not any(task_name.startswith(prefix) for prefix in _TASK_PREFIXES):
            continue
        entry: Dict[str, str] = {"task": task_name}
        if len(row) > 1:
            entry["next_run_time"] = row[1].strip()
        if len(row) > 2:
            entry["status"] = row[2].strip()
        entry["known"] = task_name in _KNOWN_TASK_NAMES
        tasks.append(entry)

    return tasks


def gather_activation_state() -> Dict[str, Any]:
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

    try:
        values = registry_tools.read_values(hive, relative_path)
    except FileNotFoundError:
        return {}
    except OSError:
        return {}

    if not values:
        return {}

    serialised: Dict[str, Any] = {}
    for key, value in values.items():
        if isinstance(value, (str, int, float, bool)) or value is None:
            serialised[str(key)] = value
        else:
            serialised[str(key)] = str(value)

    return {"path": registry_path, "values": serialised}


def gather_registry_residue() -> List[Dict[str, str]]:
    """!
    @brief Identify registry hives that likely require cleanup.
    @details The returned list mirrors OffScrub residue heuristics so planners
    can schedule deletions alongside filesystem cleanup once uninstalls complete.
    """

    entries: List[Dict[str, str]] = []

    for hive, path in _REGISTRY_RESIDUE_TEMPLATES:
        try:
            exists = registry_tools.key_exists(hive, path)
        except FileNotFoundError:
            exists = False
        except OSError:
            exists = False
        if not exists:
            continue
        entries.append({"path": _compose_handle(hive, path)})

    return entries
