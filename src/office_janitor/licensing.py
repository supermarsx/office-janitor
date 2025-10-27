"""!
@brief License and activation cleanup routines.
@details This module embeds a PowerShell script that P/Invokes the SPP/OSPP
APIs so the scrubber can remove product keys without shipping separate script
files. The orchestration coordinates registry backups, subprocess invocation,
and filesystem cleanup while respecting global safety constraints.
"""
from __future__ import annotations

import subprocess
import tempfile
from pathlib import Path
from string import Template
from typing import Iterable, Mapping, Sequence

from . import constants, fs_tools, logging_ext, registry_tools

LICENSE_SCRIPT_TEMPLATE = Template(
    r"""
function UninstallLicenses($$DllPath, $$FilterGuid) {
    $$TB = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4, 1).DefineDynamicModule(2).DefineType(0)
    [void]$$TB.DefinePInvokeMethod('SLOpen', $$DllPath, 22, 1, [int], @([IntPtr].MakeByRefType()), 1, 3)
    [void]$$TB.DefinePInvokeMethod('SLGetSLIDList', $$DllPath, 22, 1, [int], @([IntPtr], [int], [Guid].MakeByRefType(), [int], [int].MakeByRefType(), [IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
    [void]$$TB.DefinePInvokeMethod('SLUninstallLicense', $$DllPath, 22, 1, [int], @([IntPtr], [IntPtr]), 1, 3)

    $$SPPC = $$TB.CreateType()
    $$Handle = 0
    [void]$$SPPC::SLOpen([ref]$$Handle)
    $$pnReturnIds = 0
    $$ppReturnIds = 0
    $$removed = 0

    if (-not $$SPPC::SLGetSLIDList($$Handle, 0, [ref]$$FilterGuid, 6, [ref]$$pnReturnIds, [ref]$$ppReturnIds)) {
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
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    dry_run = bool(options.get("dry_run", False))
    include_spp = options.get("remove_spp", True) not in {False, "false", "0"}
    include_ospp = options.get("remove_ospp", True) not in {False, "false", "0"}
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
            "Skipping licensing cleanup because uninstall steps have not completed; use --force to override."
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
                "No backup destination configured; registry keys will not be exported before cleanup."
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
            if dry_run:
                human_logger.info("Dry-run: would execute SPP cleanup via %s", " ".join(command))
            else:
                human_logger.info("Removing SPP licenses via PowerShell script %s", script_path)
                result = subprocess.run(
                    command,
                    capture_output=True,
                    text=True,
                    timeout=300,
                    check=False,
                )
                counts = _parse_license_results(result.stdout)
                if result.returncode != 0:
                    human_logger.error(
                        "SPP cleanup failed with exit code %s", result.returncode
                    )
                    machine_logger.error(
                        "licensing_spp_failure",
                        extra={
                            "event": "licensing_spp_failure",
                            "return_code": result.returncode,
                            "stdout": result.stdout,
                            "stderr": result.stderr,
                        },
                    )
                    raise RuntimeError("SPP license removal failed")
                machine_logger.info(
                    "licensing_spp_success",
                    extra={
                        "event": "licensing_spp_success",
                        "stdout": result.stdout,
                        "stderr": result.stderr,
                        "removed": counts,
                    },
                )
        finally:
            if script_path is not None:
                try:
                    script_path.unlink(missing_ok=True)  # type: ignore[arg-type]
                except OSError:  # pragma: no cover - best effort cleanup
                    pass

    if include_ospp:
        command = options.get("ospp_command", DEFAULT_OSPP_COMMAND)
        if isinstance(command, (str, Path)):
            command = (str(command),)
        command_list = [str(part) for part in command]
        if dry_run:
            human_logger.info("Dry-run: would execute OSPP cleanup via %s", " ".join(command_list))
        else:
            human_logger.info("Removing OSPP keys via %s", command_list[0])
            result = subprocess.run(
                command_list,
                capture_output=True,
                text=True,
                timeout=300,
                check=False,
            )
            if result.returncode != 0:
                human_logger.error("OSPP cleanup failed with exit code %s", result.returncode)
                machine_logger.error(
                    "licensing_ospp_failure",
                    extra={
                        "event": "licensing_ospp_failure",
                        "return_code": result.returncode,
                        "stdout": result.stdout,
                        "stderr": result.stderr,
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

    if extra_paths:
        human_logger.info("Cleaning licensing cache directories: %s", ", ".join(map(str, extra_paths)))
        fs_tools.remove_paths(extra_paths, dry_run=dry_run)


def _render_license_script(options: Mapping[str, object]) -> str:
    """!
    @brief Build the PowerShell script used for license removal.
    @details Values from :mod:`constants` back the DLL names, GUID filters, and
    registry paths, while ``options`` allows overrides for testing or
    customisation.
    """

    guid = str(
        options.get("license_guid")
        or constants.LICENSING_GUID_FILTERS["office_family"]
    )
    spp_dll = str(options.get("spp_dll") or constants.LICENSE_DLLS["spp"])
    ospp_dll = str(options.get("ospp_dll") or constants.LICENSE_DLLS["ospp"])
    ospp_reg = str(options.get("ospp_reg_path") or constants.OSPP_REGISTRY_PATH)

    return LICENSE_SCRIPT_TEMPLATE.substitute(
        guid=guid,
        spp_dll=spp_dll,
        ospp_dll=ospp_dll,
        ospp_reg=ospp_reg,
    )


__all__ = [
    "cleanup_licenses",
]
