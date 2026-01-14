"""!
@brief License and activation cleanup routines.
@details This module embeds a PowerShell script that P/Invokes the SPP/OSPP
APIs so the scrubber can remove product keys without shipping separate script
files. The orchestration coordinates registry backups, subprocess invocation,
and filesystem cleanup while respecting global safety constraints.

Additionally provides fallback to OSPP.VBS (Microsoft's official tool) when
the PowerShell approach fails.
"""

from __future__ import annotations

import os
import tempfile
from collections.abc import Iterable, Mapping, Sequence
from pathlib import Path
from string import Template
from typing import Any

from . import constants, exec_utils, fs_tools, logging_ext, registry_tools

# Office Software Protection Platform Application ID
OFFICE_APPLICATION_ID = "0ff1ce15-a989-479d-af46-f275c6370663"
"""!
@brief The Application ID used by Office in the Software Licensing Service.
@details This GUID identifies Office licenses in SoftwareLicensingProduct WMI queries.
"""

# OSPP.VBS search locations (bundled and system-installed)
OSPP_VBS_SEARCH_PATHS: tuple[str, ...] = (
    # Bundled with office-janitor (oem folder)
    r"oem\OSPP.VBS",
    # Office 2016/2019/2021/365 (64-bit)
    r"C:\Program Files\Microsoft Office\Office16\OSPP.VBS",
    # Office 2016/2019/2021/365 (32-bit on 64-bit Windows)
    r"C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS",
    # Click-to-Run installations
    r"C:\Program Files\Microsoft Office\root\Office16\OSPP.VBS",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\OSPP.VBS",
    # Office 2013
    r"C:\Program Files\Microsoft Office\Office15\OSPP.VBS",
    r"C:\Program Files (x86)\Microsoft Office\Office15\OSPP.VBS",
)
"""!
@brief Candidate paths to locate OSPP.VBS for license management fallback.
"""

LICENSE_SCRIPT_TEMPLATE = Template(
    r"""
function UninstallLicenses($$DllPath, $$FilterGuid) {
    $$assembly = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1)
    $$module = $$assembly.DefineDynamicModule(2)
    $$TB = $$module.DefineType(0)
    [void]$$TB.DefinePInvokeMethod(
        'SLOpen', $$DllPath, 22, 1, [int], @([IntPtr].MakeByRefType()), 1, 3
    )
    [void]$$TB.DefinePInvokeMethod(
        'SLGetSLIDList',
        $$DllPath,
        22,
        1,
        [int],
        @(
            [IntPtr],
            [int],
            [Guid].MakeByRefType(),
            [int],
            [int].MakeByRefType(),
            [IntPtr].MakeByRefType()
        ),
        1,
        3
    ).SetImplementationFlags(128)
    [void]$$TB.DefinePInvokeMethod(
        'SLUninstallLicense', $$DllPath, 22, 1, [int], @([IntPtr], [IntPtr]), 1, 3
    )

    $$SPPC = $$TB.CreateType()
    $$Handle = 0
    [void]$$SPPC::SLOpen([ref]$$Handle)
    $$pnReturnIds = 0
    $$ppReturnIds = 0
    $$removed = 0

    if (-not $$SPPC::SLGetSLIDList(
        $$Handle, 0, [ref]$$FilterGuid, 6, [ref]$$pnReturnIds, [ref]$$ppReturnIds
    )) {
        if ($$pnReturnIds -gt 0) {
            foreach ($$i in 0..($$pnReturnIds - 1)) {
                [void]$$SPPC::SLUninstallLicense($$Handle, [Int64]$$ppReturnIds + ([Int64]16 * $$i))
                $$removed++
            }
        }
    }

    return $$removed
}

$$filterGuid = [Guid]"${guid}"
$$osppRegPath = "${ospp_reg}"
$$osppRoot = (Get-ItemProperty -Path $$osppRegPath -ErrorAction SilentlyContinue).Path
$$osppRemoved = 0
if ($$osppRoot) {
    $$dllPath = Join-Path $$osppRoot "${ospp_dll}"
    if (Test-Path $$dllPath) {
        $$osppRemoved = UninstallLicenses $$dllPath $$filterGuid
    }
}

$$sppRemoved = UninstallLicenses "${spp_dll}" $$filterGuid
Write-Output ("OSPP:{0}" -f $$osppRemoved)
Write-Output ("SPP:{0}" -f $$sppRemoved)
"""
)
"""!
@brief PowerShell template mirroring ``CleanOffice.txt`` with parameterised inputs.
"""

DEFAULT_OSPP_COMMAND = (
    "cscript.exe",
    "//NoLogo",
    r"C:\\Program Files\\Microsoft Office\\Office16\\ospp.vbs",
    "/unpkey:all",
)
"""!
@brief Default command tuple used to remove OSPP keys.
"""

DEFAULT_LICENSE_PATHS: tuple[Path, ...] = (
    Path(r"C:\\ProgramData\\Microsoft\\OfficeSoftwareProtectionPlatform"),
    Path(r"C:\\ProgramData\\Microsoft\\Office"),
)
"""!
@brief Common filesystem locations for Office licensing caches.
"""

DEFAULT_REGISTRY_KEYS: tuple[str, ...] = (
    constants.OSPP_REGISTRY_PATH,
    f"{constants.OSPP_REGISTRY_PATH}_Test",
)
"""!
@brief Registry keys exported prior to license cleanup.
"""


def _write_powershell_script(content: str) -> Path:
    """!
    @brief Persist a PowerShell script to disk for invocation.
    """

    with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False, encoding="utf-8") as handle:
        handle.write(content)
        return Path(handle.name)


def _expand_paths(raw: Iterable[object] | object | None) -> list[Path]:
    """!
    @brief Normalise optional path hints into :class:`Path` objects.
    """

    if raw is None:
        return []
    if isinstance(raw, (str, Path)):
        candidates = [raw]
    else:
        candidates = list(raw)
    paths: list[Path] = []
    for candidate in candidates:
        path = Path(candidate)
        if path not in paths:
            paths.append(path)
    return paths


def _expand_registry_keys(raw: Iterable[object] | object | None) -> list[str]:
    """!
    @brief Build the registry key list scheduled for export.
    """

    if raw is None:
        extras: Sequence[object] = ()
    elif isinstance(raw, (str, Path)):
        extras = [raw]
    else:
        extras = list(raw)

    keys: list[str] = list(DEFAULT_REGISTRY_KEYS)
    for entry in extras:
        if not entry:
            continue
        key = str(entry)
        if key not in keys:
            keys.append(key)
    return keys


def _resolve_backup_destination(options: Mapping[str, object]) -> Path | None:
    """!
    @brief Resolve the registry backup destination from options.
    """

    candidate = options.get("backup_destination") or options.get("backup")
    if isinstance(candidate, Path):
        return candidate
    if isinstance(candidate, str) and candidate:
        return Path(candidate)
    return None


def _parse_license_results(output: str) -> dict[str, int]:
    """!
    @brief Parse the PowerShell output for SPP/OSPP license removal counts.
    """

    counts = {"spp": 0, "ospp": 0}
    for line in output.splitlines():
        if ":" not in line:
            continue
        prefix, value = line.split(":", 1)
        prefix = prefix.strip().lower()
        try:
            parsed = int(value.strip())
        except ValueError:
            continue
        if prefix == "spp":
            counts["spp"] = parsed
        elif prefix == "ospp":
            counts["ospp"] = parsed
    return counts


def cleanup_licenses(options: Mapping[str, object]) -> None:
    """!
    @brief Remove activation artifacts based on the requested cleanup options.
    @details Accepts both the legacy option names (remove_spp, remove_ospp) and
    the new CLI flag names (clean_spp, clean_ospp, clean_vnext, clean_all_licenses)
    for backwards compatibility.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    dry_run = bool(options.get("dry_run", False))

    # Handle both legacy option names and new CLI flag names
    # Legacy: remove_spp=True by default; New CLI: clean_spp enables, clean_all_licenses enables all
    clean_all = bool(options.get("clean_all_licenses", False))
    include_spp = (
        clean_all
        or bool(options.get("clean_spp", False))
        or options.get("remove_spp", True) not in {False, "false", "0"}
    )
    include_ospp = (
        clean_all
        or bool(options.get("clean_ospp", False))
        or options.get("remove_ospp", True) not in {False, "false", "0"}
    )
    include_vnext = clean_all or bool(options.get("clean_vnext", False))

    extra_paths = _expand_paths(options.get("paths"))
    registry_keys = _expand_registry_keys(options.get("registry_keys"))
    backup_destination = _resolve_backup_destination(options)
    force_cleanup = bool(options.get("force", False))
    mode = str(options.get("mode", ""))
    uninstall_detected = bool(options.get("uninstall_detected", False))

    if not extra_paths:
        extra_paths = list(DEFAULT_LICENSE_PATHS)

    cleanup_forced = force_cleanup or mode == "cleanup-only"
    if not uninstall_detected and not cleanup_forced:
        human_logger.info(
            "Skipping licensing cleanup because uninstall steps have not completed; "
            "use --force to override."
        )
        machine_logger.info(
            "licensing_skipped",
            extra={
                "event": "licensing_skipped",
                "reason": "pending_uninstall",
                "dry_run": dry_run,
                "mode": mode,
            },
        )
        return

    machine_logger.info(
        "licensing_plan",
        extra={
            "event": "licensing_plan",
            "dry_run": dry_run,
            "remove_spp": include_spp,
            "remove_ospp": include_ospp,
            "paths": [str(path) for path in extra_paths],
            "registry_keys": registry_keys,
            "backup_destination": str(backup_destination) if backup_destination else None,
            "uninstall_detected": uninstall_detected,
            "forced": cleanup_forced,
        },
    )

    if registry_keys:
        if backup_destination is None:
            human_logger.warning(
                "No backup destination configured; registry keys will not be exported before "
                "cleanup."
            )
        else:
            if dry_run:
                human_logger.info(
                    "Dry-run: would export %d registry keys to %s.",
                    len(registry_keys),
                    backup_destination,
                )
                machine_logger.info(
                    "licensing_registry_backup",
                    extra={
                        "event": "licensing_registry_backup",
                        "keys": registry_keys,
                        "destination": str(backup_destination),
                        "dry_run": True,
                    },
                )
            else:
                try:
                    exported = registry_tools.export_keys(
                        registry_keys,
                        backup_destination,
                        dry_run=False,
                        logger=machine_logger,
                    )
                except Exception as exc:  # pragma: no cover - defensive logging
                    human_logger.warning("Failed to export registry keys prior to cleanup: %s", exc)
                    machine_logger.warning(
                        "licensing_registry_backup_failed",
                        extra={
                            "event": "licensing_registry_backup_failed",
                            "error": repr(exc),
                            "destination": str(backup_destination),
                        },
                    )
                else:
                    human_logger.info(
                        "Exported %d registry keys to %s before cleanup.",
                        len(exported),
                        backup_destination,
                    )
                    machine_logger.info(
                        "licensing_registry_backup",
                        extra={
                            "event": "licensing_registry_backup",
                            "keys": registry_keys,
                            "destination": str(backup_destination),
                            "dry_run": False,
                            "artifacts": [str(path) for path in exported],
                        },
                    )

    if include_spp:
        script_path: Path | None = None
        spp_success = False
        try:
            script_body = _render_license_script(options)
            script_path = _write_powershell_script(script_body)
            command = (
                "powershell.exe",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(script_path),
            )
            result = exec_utils.run_command(
                command,
                event="licensing_spp_cleanup",
                timeout=300,
                dry_run=dry_run,
                human_message=f"Removing SPP licenses via PowerShell script {script_path}",
                extra={"script": str(script_path)},
            )
            if not result.skipped:
                if result.returncode != 0 or result.error:
                    human_logger.warning(
                        "SPP cleanup via PowerShell failed (exit code %s); will try OSPP.VBS fallback",
                        result.returncode,
                    )
                    machine_logger.warning(
                        "licensing_spp_failure",
                        extra={
                            "event": "licensing_spp_failure",
                            "return_code": result.returncode,
                            "stdout": result.stdout,
                            "stderr": result.stderr,
                            "error": result.error,
                            "will_fallback": True,
                        },
                    )
                else:
                    counts = _parse_license_results(result.stdout)
                    spp_success = True
                    machine_logger.info(
                        "licensing_spp_success",
                        extra={
                            "event": "licensing_spp_success",
                            "stdout": result.stdout,
                            "stderr": result.stderr,
                            "removed": counts,
                        },
                    )
            else:
                spp_success = True  # Dry-run counts as success
        except Exception as exc:
            human_logger.warning("SPP cleanup exception: %s; will try OSPP.VBS fallback", exc)
            machine_logger.warning(
                "licensing_spp_exception",
                extra={
                    "event": "licensing_spp_exception",
                    "error": repr(exc),
                    "will_fallback": True,
                },
            )
        finally:
            if script_path is not None:
                try:
                    script_path.unlink(missing_ok=True)
                except OSError:  # pragma: no cover - best effort cleanup
                    pass

        # OSPP.VBS fallback if PowerShell approach failed
        if not spp_success:
            human_logger.info("Attempting license cleanup via OSPP.VBS...")
            try:
                ospp_removed = clean_licenses_via_ospp(dry_run=dry_run)
                if ospp_removed:
                    machine_logger.info(
                        "licensing_ospp_vbs_success",
                        extra={
                            "event": "licensing_ospp_vbs_success",
                            "removed_keys": ospp_removed,
                        },
                    )
                else:
                    human_logger.info("OSPP.VBS found no licenses to remove")
            except Exception as fallback_exc:
                human_logger.error("OSPP.VBS fallback also failed: %s", fallback_exc)
                machine_logger.error(
                    "licensing_ospp_vbs_failure",
                    extra={
                        "event": "licensing_ospp_vbs_failure",
                        "error": repr(fallback_exc),
                    },
                )
                raise RuntimeError("All license removal methods failed") from fallback_exc

    if include_ospp:
        command = options.get("ospp_command", DEFAULT_OSPP_COMMAND)
        if isinstance(command, (str, Path)):
            command = (str(command),)
        command_list = [str(part) for part in command]
        result = exec_utils.run_command(
            command_list,
            event="licensing_ospp_cleanup",
            timeout=300,
            dry_run=dry_run,
            human_message=f"Removing OSPP keys via {command_list[0]}",
            extra={"command": command_list},
        )
        if not result.skipped:
            if result.returncode != 0 or result.error:
                human_logger.error("OSPP cleanup failed with exit code %s", result.returncode)
                machine_logger.error(
                    "licensing_ospp_failure",
                    extra={
                        "event": "licensing_ospp_failure",
                        "return_code": result.returncode,
                        "stdout": result.stdout,
                        "stderr": result.stderr,
                        "error": result.error,
                    },
                )
                raise RuntimeError("OSPP license removal failed")
                machine_logger.info(
                    "licensing_ospp_success",
                    extra={
                        "event": "licensing_ospp_success",
                        "stdout": result.stdout,
                        "stderr": result.stderr,
                    },
                )

    # vNext identity/token cleanup
    if include_vnext:
        human_logger.info("Cleaning vNext identity cache and licensing tokens.")
        try:
            vnext_count = clean_vnext_cache(dry_run=dry_run)
            vnext_registry = registry_tools.cleanup_vnext_identity_registry(dry_run=dry_run)
            machine_logger.info(
                "licensing_vnext_success",
                extra={
                    "event": "licensing_vnext_success",
                    "cache_count": vnext_count,
                    "registry_result": vnext_registry,
                },
            )
        except Exception as exc:  # pragma: no cover - defensive logging
            human_logger.warning("vNext cleanup encountered an error: %s", exc)
            machine_logger.warning(
                "licensing_vnext_warning",
                extra={
                    "event": "licensing_vnext_warning",
                    "error": repr(exc),
                },
            )

    if extra_paths:
        human_logger.info(
            "Cleaning licensing cache directories: %s", ", ".join(map(str, extra_paths))
        )
        fs_tools.remove_paths(extra_paths, dry_run=dry_run)


def _render_license_script(options: Mapping[str, object]) -> str:
    """!
    @brief Build the PowerShell script used for license removal.
    @details Values from :mod:`constants` back the DLL names, GUID filters, and
    registry paths, while ``options`` allows overrides for testing or
    customisation.
    """

    guid = str(options.get("license_guid") or constants.LICENSING_GUID_FILTERS["office_family"])
    spp_dll = str(options.get("spp_dll") or constants.LICENSE_DLLS["spp"])
    ospp_dll = str(options.get("ospp_dll") or constants.LICENSE_DLLS["ospp"])
    ospp_reg = str(options.get("ospp_reg_path") or constants.OSPP_REGISTRY_PATH)

    return LICENSE_SCRIPT_TEMPLATE.substitute(
        guid=guid,
        spp_dll=spp_dll,
        ospp_dll=ospp_dll,
        ospp_reg=ospp_reg,
    )


def get_cleanoffice_embedded(draft_path: Path | None = None) -> str:
    """!
    @brief Return the embedded PowerShell payload from the draft `CleanOffice.txt` file.
    @param draft_path Optional path to the draft file; defaults to the repository draft location.
    @returns PowerShell script content extracted from the `:embed:` markers.
    @raises FileNotFoundError if the draft file cannot be read.
    """

    repo_draft = Path("office-janitor-draft-code") / "bin" / "CleanOffice.txt"
    path = Path(draft_path) if draft_path is not None else repo_draft
    if not path.exists():
        raise FileNotFoundError(f"Draft CleanOffice file not found: {path}")
    raw = path.read_text(encoding="utf-8")
    parts = raw.split(":embed:")
    if len(parts) < 3:
        # The payload is expected between two :embed: markers
        raise RuntimeError("Embedded payload not found in CleanOffice.txt")
    # The payload is the content between the first and second :embed:
    payload = parts[1]
    return payload


# ---------------------------------------------------------------------------
# WMI-Based License Cleanup (VBS parity: CleanOSPP subroutine)
# ---------------------------------------------------------------------------


def _query_wmi_licenses(
    application_id: str = OFFICE_APPLICATION_ID,
) -> list[dict[str, Any]]:
    """!
    @brief Query WMI for Office licenses from Software Protection Platform.
    @details Uses SoftwareLicensingProduct (Win8+) or falls back to
    OfficeSoftwareProtectionProduct (Win7).
    @param application_id The ApplicationId GUID to filter licenses.
    @returns List of license info dicts with ID, Name, PartialProductKey, ProductKeyID.
    """
    human_logger = logging_ext.get_human_logger()

    # Build WMI query via PowerShell (avoids pywin32 dependency)
    query = f"""
$results = @()
try {{
    $licenses = Get-WmiObject -Query "SELECT ID, Name, PartialProductKey, ProductKeyID FROM SoftwareLicensingProduct WHERE ApplicationId = '{application_id}' AND PartialProductKey IS NOT NULL" -ErrorAction Stop
    foreach ($lic in $licenses) {{
        $results += @{{
            ID = $lic.ID
            Name = $lic.Name
            PartialProductKey = $lic.PartialProductKey
            ProductKeyID = $lic.ProductKeyID
        }}
    }}
}} catch {{
    # Try OSPP (Win7)
    try {{
        $licenses = Get-WmiObject -Query "SELECT ID, Name, PartialProductKey, ProductKeyID FROM OfficeSoftwareProtectionProduct WHERE ApplicationId = '{application_id}' AND PartialProductKey IS NOT NULL" -ErrorAction Stop
        foreach ($lic in $licenses) {{
            $results += @{{
                ID = $lic.ID
                Name = $lic.Name
                PartialProductKey = $lic.PartialProductKey
                ProductKeyID = $lic.ProductKeyID
            }}
        }}
    }} catch {{}}
}}
$results | ConvertTo-Json -Compress
"""

    result = exec_utils.run_command(
        ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", query],
        event="query_wmi_licenses",
    )

    if result.returncode != 0 or not result.stdout:
        human_logger.debug("WMI license query returned no results or failed")
        return []

    import json

    try:
        data = json.loads(result.stdout.strip())
        if isinstance(data, dict):
            return [data]
        return list(data) if data else []
    except json.JSONDecodeError:
        human_logger.debug("Failed to parse WMI license query output")
        return []


def clean_ospp_licenses_wmi(
    *,
    dry_run: bool = False,
    application_id: str = OFFICE_APPLICATION_ID,
) -> list[str]:
    """!
    @brief Remove Office licenses from Software Protection Platform via WMI.
    @details VBS equivalent: CleanOSPP subroutine in OffScrubC2R.vbs.
    Uses SoftwareLicensingProduct.UninstallProductKey method.
    @param dry_run If True, only report what would be removed without taking action.
    @param application_id The Application GUID to filter (defaults to Office).
    @returns List of license names that were removed.
    """
    human_logger = logging_ext.get_human_logger()
    removed: list[str] = []

    licenses = _query_wmi_licenses(application_id)
    if not licenses:
        human_logger.info("No Office licenses found in Software Protection Platform")
        return removed

    human_logger.info("Found %d Office license(s) to remove", len(licenses))

    for lic in licenses:
        name = lic.get("Name", "Unknown")
        product_key_id = lic.get("ProductKeyID", "")
        partial_key = lic.get("PartialProductKey", "")

        if dry_run:
            human_logger.info(
                "[DRY-RUN] Would remove license: %s (key ending: %s)",
                name,
                partial_key,
            )
            removed.append(name)
            continue

        # Invoke UninstallProductKey via PowerShell
        uninstall_cmd = f"""
try {{
    $lic = Get-WmiObject -Query "SELECT * FROM SoftwareLicensingProduct WHERE ProductKeyID = '{product_key_id}'"
    if ($lic) {{
        $lic.UninstallProductKey($lic.ProductKeyID)
        Write-Output "OK"
    }}
}} catch {{
    try {{
        $lic = Get-WmiObject -Query "SELECT * FROM OfficeSoftwareProtectionProduct WHERE ProductKeyID = '{product_key_id}'"
        if ($lic) {{
            $lic.UninstallProductKey($lic.ProductKeyID)
            Write-Output "OK"
        }}
    }} catch {{
        Write-Output "FAIL"
    }}
}}
"""
        result = exec_utils.run_command(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", uninstall_cmd],
            event="uninstall_license_wmi",
        )

        if result.returncode == 0 and "OK" in (result.stdout or ""):
            human_logger.info("Removed license: %s (key ending: %s)", name, partial_key)
            removed.append(name)
        else:
            human_logger.warning("Failed to remove license: %s", name)

    return removed


def clean_vnext_cache(*, dry_run: bool = False) -> int:
    """!
    @brief Remove vNext license cache directories.
    @details VBS equivalent: ClearVNextLicCache subroutine in OffScrubC2R.vbs.
    Removes %LOCALAPPDATA%\\Microsoft\\Office\\Licenses and related paths.
    @param dry_run If True, only report what would be removed.
    @returns Number of paths cleaned.
    """
    import os

    human_logger = logging_ext.get_human_logger()

    cache_paths = [
        Path(os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Office\Licenses")),
        Path(os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Office\16.0\Licensing")),
        Path(os.path.expandvars(r"%PROGRAMDATA%\Microsoft\Office\Licenses")),
    ]

    existing = [p for p in cache_paths if p.exists()]
    if not existing:
        human_logger.debug("No vNext license caches found")
        return 0

    human_logger.info("Cleaning %d vNext license cache(s)", len(existing))

    if dry_run:
        for p in existing:
            human_logger.info("[DRY-RUN] Would remove: %s", p)
        return len(existing)

    fs_tools.remove_paths(existing, dry_run=False)
    return len(existing)


def clean_activation_tokens(*, dry_run: bool = False) -> int:
    """!
    @brief Remove Office activation token files.
    @details Cleans up token-based activation artifacts left behind after uninstall.
    @param dry_run If True, only report what would be removed.
    @returns Number of paths cleaned.
    """
    import os

    human_logger = logging_ext.get_human_logger()

    token_paths = [
        Path(os.path.expandvars(r"%PROGRAMDATA%\Microsoft\OfficeSoftwareProtectionPlatform")),
        Path(
            os.path.expandvars(r"%PROGRAMDATA%\Microsoft\OfficeSoftwareProtectionPlatform\Backup")
        ),
        Path(os.path.expandvars(r"%ALLUSERSPROFILE%\Microsoft\OfficeSoftwareProtectionPlatform")),
    ]

    existing = [p for p in token_paths if p.exists()]
    if not existing:
        human_logger.debug("No activation token paths found")
        return 0

    human_logger.info("Cleaning %d activation token path(s)", len(existing))

    if dry_run:
        for p in existing:
            human_logger.info("[DRY-RUN] Would remove: %s", p)
        return len(existing)

    fs_tools.remove_paths(existing, dry_run=False)
    return len(existing)


def clean_scl_cache(*, dry_run: bool = False) -> int:
    """!
    @brief Remove Shared Computer Licensing cache.
    @details Cleans up SCL token and cache directories used in shared activation scenarios.
    @param dry_run If True, only report what would be removed.
    @returns Number of paths cleaned.
    """
    import os

    human_logger = logging_ext.get_human_logger()

    scl_paths = [
        Path(os.path.expandvars(r"%PROGRAMDATA%\Microsoft\Office\SharedComputerLicensing")),
        Path(os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Office\SharedComputerLicensing")),
    ]

    existing = [p for p in scl_paths if p.exists()]
    if not existing:
        human_logger.debug("No SCL cache paths found")
        return 0

    human_logger.info("Cleaning %d SCL cache path(s)", len(existing))

    if dry_run:
        for p in existing:
            human_logger.info("[DRY-RUN] Would remove: %s", p)
        return len(existing)

    fs_tools.remove_paths(existing, dry_run=False)
    return len(existing)


# ---------------------------------------------------------------------------
# OSPP.VBS Integration (Microsoft's Official License Tool)
# ---------------------------------------------------------------------------


def find_ospp_vbs() -> Path | None:
    """!
    @brief Locate OSPP.VBS on the system.
    @details Searches bundled location first, then common Office install paths.
    @returns Path to OSPP.VBS if found, None otherwise.
    """
    # First check relative to this module (bundled version)
    module_dir = Path(__file__).parent.parent.parent  # src/office_janitor -> repo root
    bundled = module_dir / "oem" / "OSPP.VBS"
    if bundled.exists():
        return bundled

    # Check system locations
    for candidate in OSPP_VBS_SEARCH_PATHS:
        path = Path(candidate)
        if path.exists():
            return path

    return None


def _run_ospp_command(
    ospp_path: Path,
    command: str,
    *,
    dry_run: bool = False,
    timeout: int = 120,
) -> exec_utils.CommandResult:
    """!
    @brief Execute an OSPP.VBS command via cscript.
    @param ospp_path Path to OSPP.VBS.
    @param command The OSPP command (e.g., "/dstatus", "/unpkey:XXXXX").
    @param dry_run If True, log but don't execute.
    @param timeout Seconds to wait for command completion.
    @returns CommandResult with stdout, stderr, and return code.
    """
    cmd = [
        "cscript.exe",
        "//NoLogo",
        str(ospp_path),
        command,
    ]

    return exec_utils.run_command(
        cmd,
        event="ospp_vbs",
        timeout=timeout,
        dry_run=dry_run,
        human_message=f"Running OSPP.VBS {command}",
        extra={"ospp_path": str(ospp_path), "command": command},
    )


def query_ospp_status(ospp_path: Path | None = None) -> list[dict[str, str]]:
    """!
    @brief Query license status using OSPP.VBS /dstatus.
    @param ospp_path Path to OSPP.VBS, or None to auto-detect.
    @returns List of license info dicts with Name, PartialProductKey, Status.
    """
    human_logger = logging_ext.get_human_logger()

    if ospp_path is None:
        ospp_path = find_ospp_vbs()
        if ospp_path is None:
            human_logger.debug("OSPP.VBS not found on system")
            return []

    result = _run_ospp_command(ospp_path, "/dstatus")
    if result.returncode != 0 or not result.stdout:
        human_logger.debug("OSPP.VBS /dstatus failed or returned no output")
        return []

    return _parse_ospp_dstatus(result.stdout)


def _parse_ospp_dstatus(output: str) -> list[dict[str, str]]:
    """!
    @brief Parse OSPP.VBS /dstatus output into structured data.
    @param output Raw stdout from /dstatus command.
    @returns List of license dicts with parsed fields.
    """
    licenses: list[dict[str, str]] = []
    current: dict[str, str] = {}

    for line in output.splitlines():
        line = line.strip()
        if not line:
            continue

        # Each license block starts with LICENSE NAME:
        if line.startswith("LICENSE NAME:"):
            if current:
                licenses.append(current)
            current = {"name": line.split(":", 1)[1].strip()}
        elif line.startswith("LICENSE DESCRIPTION:"):
            current["description"] = line.split(":", 1)[1].strip()
        elif line.startswith("LICENSE STATUS:"):
            current["status"] = line.split(":", 1)[1].strip()
        elif line.startswith("Last 5 characters of installed product key:"):
            current["partial_key"] = line.split(":", 1)[1].strip()
        elif line.startswith("SKU ID:"):
            current["sku_id"] = line.split(":", 1)[1].strip()
        elif line.startswith("ERROR CODE:"):
            current["error_code"] = line.split(":", 1)[1].strip()

    if current:
        licenses.append(current)

    return licenses


def uninstall_ospp_key(
    partial_key: str,
    ospp_path: Path | None = None,
    *,
    dry_run: bool = False,
) -> bool:
    """!
    @brief Uninstall a product key using OSPP.VBS /unpkey.
    @param partial_key The last 5 characters of the product key.
    @param ospp_path Path to OSPP.VBS, or None to auto-detect.
    @param dry_run If True, log but don't execute.
    @returns True if key was uninstalled successfully.
    """
    human_logger = logging_ext.get_human_logger()

    if len(partial_key) != 5:
        human_logger.warning("Invalid partial key length: %s (expected 5 characters)", partial_key)
        return False

    if ospp_path is None:
        ospp_path = find_ospp_vbs()
        if ospp_path is None:
            human_logger.warning("OSPP.VBS not found; cannot uninstall key")
            return False

    result = _run_ospp_command(ospp_path, f"/unpkey:{partial_key}", dry_run=dry_run)

    if result.skipped:
        return True  # Dry-run counts as success

    # Check for success messages in output
    stdout = result.stdout or ""
    if (
        "Product key uninstall successful" in stdout
        or "<Product key uninstall successful>" in stdout
    ):
        human_logger.info("Successfully uninstalled product key ending in %s", partial_key)
        return True

    if "Product key not found" in stdout:
        human_logger.debug("Product key %s not found (may already be removed)", partial_key)
        return True  # Not an error if already gone

    human_logger.warning("Failed to uninstall product key %s: %s", partial_key, stdout)
    return False


def clean_licenses_via_ospp(
    ospp_path: Path | None = None,
    *,
    dry_run: bool = False,
) -> list[str]:
    """!
    @brief Remove all Office licenses using OSPP.VBS as fallback.
    @details Queries /dstatus to find all installed keys, then uses /unpkey
    to remove each one. This is the official Microsoft approach.
    @param ospp_path Path to OSPP.VBS, or None to auto-detect.
    @param dry_run If True, log but don't execute.
    @returns List of partial keys that were removed.
    """
    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    if ospp_path is None:
        ospp_path = find_ospp_vbs()
        if ospp_path is None:
            human_logger.warning("OSPP.VBS not found; skipping OSPP-based license cleanup")
            return []

    human_logger.info("Querying license status via OSPP.VBS...")
    licenses = query_ospp_status(ospp_path)

    if not licenses:
        human_logger.info("No Office licenses found via OSPP.VBS")
        return []

    # Filter to licenses that have a partial key (can be uninstalled)
    removable = [lic for lic in licenses if lic.get("partial_key")]
    if not removable:
        human_logger.info("No removable licenses found (no partial keys)")
        return []

    human_logger.info("Found %d Office license(s) to remove via OSPP.VBS", len(removable))
    machine_logger.info(
        "ospp_licenses_found",
        extra={
            "event": "ospp_licenses_found",
            "count": len(removable),
            "licenses": [
                {"name": lic.get("name", ""), "partial_key": lic.get("partial_key", "")}
                for lic in removable
            ],
        },
    )

    removed: list[str] = []
    for lic in removable:
        partial_key = lic.get("partial_key", "")
        name = lic.get("name", "Unknown")

        if dry_run:
            human_logger.info("[DRY-RUN] Would remove license: %s (key: %s)", name, partial_key)
            removed.append(partial_key)
            continue

        if uninstall_ospp_key(partial_key, ospp_path, dry_run=False):
            removed.append(partial_key)
            machine_logger.info(
                "ospp_license_removed",
                extra={
                    "event": "ospp_license_removed",
                    "name": name,
                    "partial_key": partial_key,
                },
            )
        else:
            machine_logger.warning(
                "ospp_license_removal_failed",
                extra={
                    "event": "ospp_license_removal_failed",
                    "name": name,
                    "partial_key": partial_key,
                },
            )

    return removed


def full_license_cleanup(
    *,
    dry_run: bool = False,
    keep_license: bool = False,
    use_ospp_fallback: bool = True,
) -> dict[str, Any]:
    """!
    @brief Complete Office license cleanup combining all license removal methods.
    @details VBS parity: CleanOSPP + ClearVNextLicCache combined.
    Falls back to OSPP.VBS if WMI-based cleanup fails or finds no licenses.
    @param dry_run If True, only report what would be removed.
    @param keep_license If True, skip license cleanup entirely.
    @param use_ospp_fallback If True, use OSPP.VBS when WMI approach fails.
    @returns Dict with counts of cleaned items per category.
    """
    human_logger = logging_ext.get_human_logger()

    if keep_license:
        return {"skipped": True, "reason": "keep_license flag set"}

    results: dict[str, Any] = {}

    # Try WMI-based cleanup first
    try:
        results["ospp_wmi"] = clean_ospp_licenses_wmi(dry_run=dry_run)
    except Exception as exc:
        human_logger.warning("WMI-based license cleanup failed: %s", exc)
        results["ospp_wmi"] = []
        results["ospp_wmi_error"] = str(exc)

    # Fallback to OSPP.VBS if WMI found nothing or failed
    if use_ospp_fallback and not results.get("ospp_wmi"):
        human_logger.info("Attempting license cleanup via OSPP.VBS fallback...")
        try:
            results["ospp_vbs"] = clean_licenses_via_ospp(dry_run=dry_run)
        except Exception as exc:
            human_logger.warning("OSPP.VBS fallback failed: %s", exc)
            results["ospp_vbs_error"] = str(exc)

    results["vnext_cache"] = clean_vnext_cache(dry_run=dry_run)
    results["activation_tokens"] = clean_activation_tokens(dry_run=dry_run)
    results["scl_cache"] = clean_scl_cache(dry_run=dry_run)

    return results


__all__ = [
    "cleanup_licenses",
    "get_cleanoffice_embedded",
    "OFFICE_APPLICATION_ID",
    "OSPP_VBS_SEARCH_PATHS",
    "find_ospp_vbs",
    "query_ospp_status",
    "uninstall_ospp_key",
    "clean_licenses_via_ospp",
    "clean_ospp_licenses_wmi",
    "clean_vnext_cache",
    "clean_activation_tokens",
    "clean_scl_cache",
    "full_license_cleanup",
]
