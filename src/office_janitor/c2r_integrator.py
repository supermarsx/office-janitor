"""!
@brief C2R integrator.exe operations for unregistration and license management.
@details Provides utilities for interacting with the Office C2R integrator.exe
executable to unregister integration components and reinstall licenses.
Based on functionality from OffScrubC2R.vbs and OfficeScrubber.cmd.
"""

from __future__ import annotations

import glob
import os
import winreg
from pathlib import Path

from . import command_runner, exec_utils, logging_ext

# ---------------------------------------------------------------------------
# Integrator.exe Locations
# ---------------------------------------------------------------------------

INTEGRATOR_EXE_CANDIDATES = (
    Path(r"C:\Program Files\Common Files\Microsoft Shared\ClickToRun\integrator.exe"),
    Path(r"C:\Program Files (x86)\Common Files\Microsoft Shared\ClickToRun\integrator.exe"),
)
"""!
@brief Default locations for the C2R integrator executable.
"""

INTEGRATOR_TIMEOUT = 120
"""!
@brief Default timeout for integrator.exe operations in seconds.
"""

# ---------------------------------------------------------------------------
# Integrator.exe Lookup
# ---------------------------------------------------------------------------


def find_integrator_exe() -> Path | None:
    """!
    @brief Locate the C2R integrator.exe binary.
    @returns Path to integrator.exe or None if not found.
    """
    for candidate in INTEGRATOR_EXE_CANDIDATES:
        if candidate.exists():
            return candidate
    return None


def find_integrator_in_package(package_folder: Path) -> Path | None:
    """!
    @brief Locate integrator.exe within a C2R package folder.
    @param package_folder Root folder of the C2R package.
    @returns Path to integrator.exe or None if not found.
    """
    candidates = [
        package_folder
        / "root"
        / "vfs"
        / "ProgramFilesCommonX64"
        / "Microsoft Shared"
        / "ClickToRun"
        / "integrator.exe",
        package_folder / "root" / "Integration" / "integrator.exe",
        package_folder / "Integration" / "integrator.exe",
    ]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return None


# ---------------------------------------------------------------------------
# Manifest Cleanup
# ---------------------------------------------------------------------------


def delete_c2r_manifests(
    package_folder: Path | str,
    *,
    dry_run: bool = False,
) -> list[Path]:
    """!
    @brief Delete C2RManifest*.xml files from a package's Integration folder.
    @details VBS equivalent: del command for C2RManifest*.xml in OffScrubC2R.vbs.
    @param package_folder Root folder of the C2R package
        (e.g., C:\\Program Files\\Microsoft Office).
    @param dry_run If True, only log what would be deleted.
    @returns List of manifest files deleted (or that would be deleted in dry-run).
    """
    human_logger = logging_ext.get_human_logger()
    package_folder = Path(package_folder)

    integration_path = package_folder / "root" / "Integration"
    if not integration_path.exists():
        human_logger.debug("Integration folder not found: %s", integration_path)
        return []

    manifest_pattern = str(integration_path / "C2RManifest*.xml")
    manifest_files = [Path(p) for p in glob.glob(manifest_pattern)]

    deleted: list[Path] = []
    for manifest in manifest_files:
        if dry_run:
            human_logger.info("[DRY-RUN] Would delete manifest: %s", manifest)
            deleted.append(manifest)
        else:
            try:
                manifest.unlink()
                human_logger.debug("Deleted manifest: %s", manifest)
                deleted.append(manifest)
            except OSError as exc:
                human_logger.warning("Failed to delete manifest %s: %s", manifest, exc)

    return deleted


# ---------------------------------------------------------------------------
# C2R Unregistration
# ---------------------------------------------------------------------------


def unregister_c2r_integration(
    package_folder: Path | str,
    package_guid: str,
    *,
    dry_run: bool = False,
    timeout: int = INTEGRATOR_TIMEOUT,
) -> int:
    """!
    @brief Unregister C2R integration components via integrator.exe.
    @details VBS equivalent: integrator.exe /U /Extension call in OffScrubC2R.vbs.
        Steps:
        1. Delete C2RManifest*.xml files from the Integration folder
        2. Call integrator.exe /U /Extension with PackageRoot and PackageGUID
    @param package_folder Root folder of the C2R package.
    @param package_guid The PackageGUID for unregistration.
    @param dry_run If True, only log what would be done.
    @param timeout Timeout for integrator.exe command.
    @returns Exit code from integrator.exe (0 = success, -1 if not found).
    """
    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()
    package_folder = Path(package_folder)

    # Step 1: Delete manifest files
    deleted_manifests = delete_c2r_manifests(package_folder, dry_run=dry_run)
    if deleted_manifests:
        human_logger.info("Deleted %d C2R manifest file(s)", len(deleted_manifests))

    # Step 2: Find integrator.exe
    integrator = find_integrator_exe()
    if integrator is None:
        integrator = find_integrator_in_package(package_folder)

    if integrator is None:
        human_logger.debug("Integrator.exe not found, skipping unregistration")
        return -1

    # Step 3: Build and execute unregister command
    # Format: integrator.exe /U /Extension PackageRoot=<path> PackageGUID=<guid>
    command = [
        str(integrator),
        "/U",
        "/Extension",
        f"PackageRoot={package_folder}",
        f"PackageGUID={package_guid}",
    ]

    if dry_run:
        human_logger.info("[DRY-RUN] Would execute: %s", " ".join(command))
        return 0

    human_logger.info("Unregistering C2R integration for: %s", package_folder)
    machine_logger.info(
        "c2r_unregister_start",
        extra={
            "event": "c2r_unregister_start",
            "package_folder": str(package_folder),
            "package_guid": package_guid,
        },
    )

    result = command_runner.run_command(
        command,
        timeout=timeout,
        event="c2r_unregister",
    )

    machine_logger.info(
        "c2r_unregister_complete",
        extra={
            "event": "c2r_unregister_complete",
            "exit_code": result.returncode,
        },
    )

    if result.returncode == 0:
        human_logger.debug("C2R unregistration completed successfully")
    else:
        human_logger.warning("C2R unregistration exited with code: %d", result.returncode)

    return result.returncode


def find_c2r_package_guids() -> list[tuple[Path, str]]:
    """!
    @brief Find installed C2R package folders and their GUIDs from registry.
    @details Scans HKLM\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration for
        PackageGUID and installation path information.
    @returns List of (package_folder, package_guid) tuples.
    """
    human_logger = logging_ext.get_human_logger()
    results: list[tuple[Path, str]] = []

    config_keys = [
        r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
        r"SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration",
    ]

    for key_path in config_keys:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_READ) as key:
                try:
                    package_guid = winreg.QueryValueEx(key, "PackageGUID")[0]
                except FileNotFoundError:
                    continue

                # Try to get install path
                install_path = None
                for value_name in ("InstallationPath", "ClientFolder"):
                    try:
                        install_path = winreg.QueryValueEx(key, value_name)[0]
                        break
                    except FileNotFoundError:
                        continue

                if install_path and package_guid:
                    results.append((Path(install_path), package_guid))
                    human_logger.debug("Found C2R package: %s (%s)", install_path, package_guid)

        except FileNotFoundError:
            continue
        except OSError as exc:
            human_logger.debug("Failed to read %s: %s", key_path, exc)
            continue

    return results


def unregister_all_c2r_integrations(*, dry_run: bool = False) -> int:
    """!
    @brief Unregister all found C2R integration components.
    @details Discovers all C2R packages from registry and unregisters each.
    @param dry_run If True, only log what would be done.
    @returns Number of packages successfully unregistered.
    """
    human_logger = logging_ext.get_human_logger()

    packages = find_c2r_package_guids()
    if not packages:
        human_logger.debug("No C2R packages found for unregistration")
        return 0

    success_count = 0
    for package_folder, package_guid in packages:
        result = unregister_c2r_integration(
            package_folder,
            package_guid,
            dry_run=dry_run,
        )
        if result == 0:
            success_count += 1

    human_logger.info("Unregistered %d of %d C2R packages", success_count, len(packages))
    return success_count


# ---------------------------------------------------------------------------
# License Management
# ---------------------------------------------------------------------------


def get_c2r_product_release_ids() -> list[str]:
    """!
    @brief Get Office C2R product release IDs (SKUs) from registry.
    @details Scans ProductReleaseIDs under the active configuration to find
        installed Office SKUs like ProPlus, Professional, Standard, etc.
    @returns List of SKU names (e.g., ["ProPlus2024Retail", "VisioProRetail"]).
    """
    human_logger = logging_ext.get_human_logger()
    skus: list[str] = []

    prids_keys = [
        r"SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs",
        r"SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\ProductReleaseIDs",
    ]

    for key_path in prids_keys:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_READ) as key:
                # Get ActiveConfiguration to find the channel
                try:
                    active_config = winreg.QueryValueEx(key, "ActiveConfiguration")[0]
                except FileNotFoundError:
                    continue

                # Open the configuration subkey
                config_path = f"{key_path}\\{active_config}"
                try:
                    with winreg.OpenKey(
                        winreg.HKEY_LOCAL_MACHINE, config_path, 0, winreg.KEY_READ
                    ) as config_key:
                        # Enumerate subkeys to find product SKUs
                        idx = 0
                        while True:
                            try:
                                subkey_name = winreg.EnumKey(config_key, idx)
                                # SKUs have ".16" suffix - extract just the name
                                if subkey_name.endswith(".16"):
                                    sku_name = subkey_name[:-3]  # Remove ".16"
                                    skus.append(sku_name)
                                    human_logger.debug("Found SKU: %s", sku_name)
                                idx += 1
                            except OSError:
                                break
                except FileNotFoundError:
                    continue

        except FileNotFoundError:
            continue
        except OSError as exc:
            human_logger.debug("Failed to read %s: %s", key_path, exc)
            continue

    return skus


def get_c2r_install_root() -> tuple[Path | None, str | None]:
    """!
    @brief Get C2R install root path and package GUID.
    @details Reads InstallPath and PackageGUID from registry, appending '\\root'
        to InstallPath as expected by integrator.exe /R /License.
    @returns Tuple of (install_root_path, package_guid) or (None, None) if not found.
    """
    human_logger = logging_ext.get_human_logger()

    config_keys = [
        r"SOFTWARE\Microsoft\Office\ClickToRun",
        r"SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun",
    ]

    for key_path in config_keys:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_READ) as key:
                try:
                    install_path = winreg.QueryValueEx(key, "InstallPath")[0]
                    package_guid = winreg.QueryValueEx(key, "PackageGUID")[0]

                    # Add \root as per OfficeScrubber.cmd logic
                    install_root = Path(install_path) / "root"
                    human_logger.debug(
                        "Found C2R install root: %s (GUID: %s)", install_root, package_guid
                    )
                    return install_root, package_guid
                except FileNotFoundError:
                    continue
        except FileNotFoundError:
            continue
        except OSError as exc:
            human_logger.debug("Failed to read %s: %s", key_path, exc)
            continue

    return None, None


def reinstall_c2r_license(
    sku_name: str,
    package_root: Path | str,
    package_guid: str,
    *,
    dry_run: bool = False,
    timeout: int = INTEGRATOR_TIMEOUT,
) -> int:
    """!
    @brief Reinstall Office C2R license for a single product SKU.
    @details Calls integrator.exe /R /License to reinstall license files.
        Based on OfficeScrubber.cmd license reset functionality (option T).
    @param sku_name Product SKU name (e.g., "ProPlus2024Retail").
    @param package_root C2R package root path (with \\root suffix).
    @param package_guid C2R package GUID.
    @param dry_run If True, only log what would be done.
    @param timeout Timeout for integrator.exe command.
    @returns Exit code from integrator.exe (0 = success).
    """
    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    # Find integrator.exe
    integrator = find_integrator_exe()
    if integrator is None:
        # Check within package folder
        pkg_integrator = Path(package_root) / "Integration" / "integrator.exe"
        if pkg_integrator.exists():
            integrator = pkg_integrator
        else:
            human_logger.warning("Integrator.exe not found, cannot reinstall license")
            return -1

    # Build command:
    # integrator.exe /R /License PRIDName=<sku>.16 PackageGUID=<guid> PackageRoot=<path>
    prid_name = f"{sku_name}.16"
    command = [
        str(integrator),
        "/R",
        "/License",
        f"PRIDName={prid_name}",
        f"PackageGUID={package_guid}",
        f"PackageRoot={package_root}",
    ]

    if dry_run:
        human_logger.info("[DRY-RUN] Would reinstall license: %s", prid_name)
        return 0

    human_logger.info("Reinstalling license for: %s", sku_name)
    machine_logger.info(
        "c2r_license_reinstall_start",
        extra={
            "event": "c2r_license_reinstall_start",
            "sku_name": sku_name,
            "prid_name": prid_name,
            "package_root": str(package_root),
            "package_guid": package_guid,
        },
    )

    result = command_runner.run_command(
        command,
        timeout=timeout,
        event="c2r_license_reinstall",
    )

    machine_logger.info(
        "c2r_license_reinstall_complete",
        extra={
            "event": "c2r_license_reinstall_complete",
            "sku_name": sku_name,
            "exit_code": result.returncode,
        },
    )

    if result.returncode == 0:
        human_logger.debug("License reinstall completed for %s", sku_name)
    else:
        human_logger.warning(
            "License reinstall for %s exited with code: %d", sku_name, result.returncode
        )

    return result.returncode


def reinstall_c2r_licenses(
    *, dry_run: bool = False, timeout: int = INTEGRATOR_TIMEOUT
) -> dict[str, int]:
    """!
    @brief Reinstall all Office C2R licenses using integrator.exe.
    @details Resets Office licensing by reinstalling license files for all
        detected product SKUs. Based on OfficeScrubber.cmd license menu option T.
        Steps:
        1. Detect installed C2R configuration (InstallPath, PackageGUID)
        2. Enumerate product SKUs from ProductReleaseIDs
        3. Call integrator.exe /R /License for each SKU
    @param dry_run If True, only log what would be done.
    @param timeout Timeout for each integrator.exe command.
    @returns Dictionary mapping SKU names to exit codes.
    """
    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    # Step 1: Get install root and package GUID
    install_root, package_guid = get_c2r_install_root()
    if install_root is None or package_guid is None:
        human_logger.warning("No installed Office C2R detected, cannot reinstall licenses")
        return {}

    # Check integrator.exe exists
    integrator = find_integrator_exe()
    if integrator is None:
        pkg_integrator = install_root / "Integration" / "integrator.exe"
        if not pkg_integrator.exists():
            human_logger.warning("Integrator.exe not found, cannot reinstall licenses")
            return {}

    # Step 2: Get product SKUs
    skus = get_c2r_product_release_ids()
    if not skus:
        human_logger.warning("No product SKUs found, cannot reinstall licenses")
        return {}

    human_logger.info("Found %d Office product SKU(s) for license reinstall", len(skus))
    machine_logger.info(
        "c2r_licenses_reinstall_start",
        extra={
            "event": "c2r_licenses_reinstall_start",
            "sku_count": len(skus),
            "skus": skus,
            "package_root": str(install_root),
            "package_guid": package_guid,
            "dry_run": dry_run,
        },
    )

    # Step 3: Reinstall each SKU
    results: dict[str, int] = {}
    for sku in skus:
        exit_code = reinstall_c2r_license(
            sku,
            install_root,
            package_guid,
            dry_run=dry_run,
            timeout=timeout,
        )
        results[sku] = exit_code

    successes = sum(1 for code in results.values() if code == 0)
    failures = len(results) - successes

    human_logger.info("License reinstall complete: %d succeeded, %d failed", successes, failures)
    machine_logger.info(
        "c2r_licenses_reinstall_complete",
        extra={
            "event": "c2r_licenses_reinstall_complete",
            "successes": successes,
            "failures": failures,
            "results": results,
        },
    )

    return results


# ---------------------------------------------------------------------------
# C2R Update Channel definitions
# ---------------------------------------------------------------------------

C2R_UPDATE_CHANNELS: dict[str, tuple[str, str]] = {
    "current": (
        "Current Channel",
        "http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60",
    ),
    "monthly": (
        "Monthly Enterprise Channel",
        "http://officecdn.microsoft.com/pr/55336b82-a18d-4dd6-b5f6-9e5095c314a6",
    ),
    "semi-annual": (
        "Semi-Annual Enterprise Channel",
        "http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114",
    ),
    "beta": (
        "Beta Channel",
        "http://officecdn.microsoft.com/pr/5440fd1f-7ecb-4221-8110-145efaa6372f",
    ),
    "current-preview": (
        "Current Channel (Preview)",
        "http://officecdn.microsoft.com/pr/64256afe-f5d9-4f86-8936-8840a6a4f5be",
    ),
    "semi-annual-preview": (
        "Semi-Annual Enterprise Channel (Preview)",
        "http://officecdn.microsoft.com/pr/b8f9b850-328d-4355-9145-c59439a0c4cf",
    ),
}


def trigger_update(*, dry_run: bool = False) -> dict[str, object]:
    """!
    @brief Trigger an Office C2R update check.
    @param dry_run If True, don't actually run the command.
    @returns Dict with success status and any error message.
    """
    human_logger = logging_ext.get_human_logger()

    # Find OfficeC2RClient.exe
    c2r_client = _find_office_c2r_client()
    if c2r_client is None:
        return {"success": False, "error": "OfficeC2RClient.exe not found"}

    if dry_run:
        human_logger.info("[DRY-RUN] Would run: OfficeC2RClient.exe /update user")
        return {"success": True, "dry_run": True}

    cmd = [str(c2r_client), "/update", "user"]
    result = exec_utils.run_command(
        cmd,
        event="c2r_update",
        timeout=60,
        dry_run=False,
        human_message="Triggering Office update check",
    )

    if result.returncode == 0:
        human_logger.info("Office update check triggered successfully")
        return {"success": True}
    else:
        error_msg = result.stderr or result.stdout or "Unknown error"
        return {"success": False, "error": error_msg}


def change_update_channel(
    channel_id: str,
    *,
    dry_run: bool = False,
) -> dict[str, object]:
    """!
    @brief Change the Office C2R update channel.
    @param channel_id Key from C2R_UPDATE_CHANNELS (e.g., 'current', 'monthly').
    @param dry_run If True, don't actually make changes.
    @returns Dict with success status and any error message.
    """
    human_logger = logging_ext.get_human_logger()

    if channel_id not in C2R_UPDATE_CHANNELS:
        return {"success": False, "error": f"Unknown channel: {channel_id}"}

    channel_name, cdn_url = C2R_UPDATE_CHANNELS[channel_id]

    # Find OfficeC2RClient.exe
    c2r_client = _find_office_c2r_client()
    if c2r_client is None:
        return {"success": False, "error": "OfficeC2RClient.exe not found"}

    if dry_run:
        human_logger.info(f"[DRY-RUN] Would change channel to: {channel_name}")
        return {"success": True, "dry_run": True, "channel": channel_name}

    # Use /changesetting to update the CDN URL
    cmd = [str(c2r_client), "/changesetting", f"CDNBaseUrl={cdn_url}"]
    result = exec_utils.run_command(
        cmd,
        event="c2r_channel_change",
        timeout=60,
        dry_run=False,
        human_message=f"Changing update channel to {channel_name}",
    )

    if result.returncode == 0:
        human_logger.info(f"Office update channel changed to {channel_name}")
        return {"success": True, "channel": channel_name}
    else:
        error_msg = result.stderr or result.stdout or "Unknown error"
        return {"success": False, "error": error_msg}


def _find_office_c2r_client() -> Path | None:
    """!
    @brief Find OfficeC2RClient.exe in common locations.
    @returns Path to OfficeC2RClient.exe or None if not found.
    """
    search_paths = [
        Path(os.environ.get("ProgramFiles", r"C:\Program Files"))
        / "Common Files"
        / "Microsoft Shared"
        / "ClickToRun"
        / "OfficeC2RClient.exe",
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)"))
        / "Common Files"
        / "Microsoft Shared"
        / "ClickToRun"
        / "OfficeC2RClient.exe",
    ]

    for path in search_paths:
        if path.exists():
            return path

    return None
