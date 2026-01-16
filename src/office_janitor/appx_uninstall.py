"""!
@file appx_uninstall.py
@brief Microsoft Store (AppX) Office package removal utilities.
@details Provides functions to detect and remove Office packages installed via
    the Microsoft Store using PowerShell AppX cmdlets.
"""

from __future__ import annotations

import logging
import subprocess
from typing import TYPE_CHECKING

from . import constants

if TYPE_CHECKING:
    from collections.abc import Sequence

__all__ = [
    "detect_office_appx_packages",
    "remove_office_appx_packages",
    "remove_provisioned_appx_packages",
    "get_appx_package_info",
]

_logger = logging.getLogger(__name__)


def _run_powershell(
    command: str,
    *,
    capture_output: bool = True,
    timeout: int = 120,
) -> subprocess.CompletedProcess[str]:
    """!
    @brief Execute a PowerShell command.
    @param command The PowerShell command to execute.
    @param capture_output Whether to capture stdout/stderr.
    @param timeout Timeout in seconds.
    @returns CompletedProcess result.
    """
    return subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", command],
        capture_output=capture_output,
        text=True,
        timeout=timeout,
        check=False,
    )


def detect_office_appx_packages() -> list[dict[str, str]]:
    """!
    @brief Detect installed Microsoft Store Office packages.
    @returns List of dictionaries with package info (Name, PackageFullName, Version).
    """
    packages: list[dict[str, str]] = []

    for pattern in constants.OFFICE_APPX_PACKAGES:
        try:
            # Query for packages matching the pattern
            cmd = (
                f'Get-AppxPackage -Name "*{pattern}*" -AllUsers 2>$null | '
                f"Select-Object Name, PackageFullName, Version | "
                f"ConvertTo-Json -Compress"
            )
            result = _run_powershell(cmd, timeout=60)

            if result.returncode == 0 and result.stdout.strip():
                import json

                try:
                    data = json.loads(result.stdout)
                    # Handle single result (dict) vs multiple results (list)
                    if isinstance(data, dict):
                        packages.append(data)
                    elif isinstance(data, list):
                        packages.extend(data)
                except json.JSONDecodeError:
                    _logger.debug("Failed to parse AppX query result for %s", pattern)

        except subprocess.TimeoutExpired:
            _logger.warning("Timeout querying AppX packages for pattern: %s", pattern)
        except OSError as e:
            _logger.warning("Failed to query AppX packages: %s", e)

    # Deduplicate by PackageFullName
    seen: set[str] = set()
    unique: list[dict[str, str]] = []
    for pkg in packages:
        full_name = pkg.get("PackageFullName", "")
        if full_name and full_name not in seen:
            seen.add(full_name)
            unique.append(pkg)

    return unique


def get_appx_package_info(package_name: str) -> dict[str, str] | None:
    """!
    @brief Get detailed info for a specific AppX package.
    @param package_name The package name or full name.
    @returns Dictionary with package info or None if not found.
    """
    try:
        cmd = (
            f'Get-AppxPackage -Name "*{package_name}*" -AllUsers 2>$null | '
            f"Select-Object Name, PackageFullName, Version, InstallLocation, "
            f"Publisher, Architecture | ConvertTo-Json -Compress"
        )
        result = _run_powershell(cmd, timeout=30)

        if result.returncode == 0 and result.stdout.strip():
            import json

            try:
                data = json.loads(result.stdout)
                if isinstance(data, list) and data:
                    return data[0]
                if isinstance(data, dict):
                    return data
            except json.JSONDecodeError:
                pass
    except (subprocess.TimeoutExpired, OSError) as e:
        _logger.debug("Failed to get AppX package info for %s: %s", package_name, e)

    return None


def remove_office_appx_packages(
    packages: Sequence[str] | None = None,
    *,
    dry_run: bool = False,
    all_users: bool = True,
) -> list[dict[str, object]]:
    """!
    @brief Remove Microsoft Store Office packages.
    @param packages List of package names to remove, or None to remove all Office packages.
    @param dry_run If True, only log what would be removed without actually removing.
    @param all_users If True, remove for all users (requires admin).
    @returns List of results with package name and success status.
    """
    if packages is None:
        # Detect all installed Office AppX packages
        detected = detect_office_appx_packages()
        packages = [pkg.get("PackageFullName", pkg.get("Name", "")) for pkg in detected]

    if not packages:
        _logger.info("No Microsoft Store Office packages found to remove")
        return []

    results: list[dict[str, object]] = []
    all_users_flag = "-AllUsers" if all_users else ""

    for package in packages:
        if not package:
            continue

        result: dict[str, object] = {
            "package": package,
            "success": False,
            "dry_run": dry_run,
            "error": None,
        }

        if dry_run:
            _logger.info("[DRY-RUN] Would remove AppX package: %s", package)
            result["success"] = True
            results.append(result)
            continue

        try:
            # Use PackageFullName if available, otherwise search by name
            if "{" in package or "_" in package:
                # Looks like a full package name
                cmd = f'Get-AppxPackage "{package}" {all_users_flag} | Remove-AppxPackage'
            else:
                # Search by name pattern
                cmd = f'Get-AppxPackage -Name "*{package}*" {all_users_flag} | Remove-AppxPackage'

            _logger.info("Removing AppX package: %s", package)
            proc_result = _run_powershell(cmd, timeout=300)

            if proc_result.returncode == 0:
                result["success"] = True
                _logger.info("Successfully removed AppX package: %s", package)
            else:
                result["error"] = proc_result.stderr.strip() or "Unknown error"
                _logger.warning("Failed to remove AppX package %s: %s", package, result["error"])

        except subprocess.TimeoutExpired:
            result["error"] = "Timeout"
            _logger.warning("Timeout removing AppX package: %s", package)
        except OSError as e:
            result["error"] = str(e)
            _logger.warning("Error removing AppX package %s: %s", package, e)

        results.append(result)

    return results


def remove_provisioned_appx_packages(
    *,
    dry_run: bool = False,
) -> list[dict[str, object]]:
    """!
    @brief Remove provisioned (system-wide) Office AppX packages.
    @details Provisioned packages are automatically installed for new users.
        Removing them prevents Office from being installed for future users.
    @param dry_run If True, only log what would be removed.
    @returns List of results with package name and success status.
    """
    results: list[dict[str, object]] = []

    for pattern in constants.OFFICE_APPX_PROVISIONED_PACKAGES:
        try:
            # Find provisioned packages matching pattern
            find_cmd = (
                f"Get-AppxProvisionedPackage -Online 2>$null | "
                f"Where-Object {{ $_.DisplayName -like '{pattern}' }} | "
                f"Select-Object DisplayName, PackageName | ConvertTo-Json -Compress"
            )
            find_result = _run_powershell(find_cmd, timeout=60)

            if find_result.returncode != 0 or not find_result.stdout.strip():
                continue

            import json

            try:
                data = json.loads(find_result.stdout)
                if isinstance(data, dict):
                    data = [data]
            except json.JSONDecodeError:
                continue

            for pkg in data:
                pkg_name = pkg.get("PackageName", "")
                display_name = pkg.get("DisplayName", pkg_name)

                if not pkg_name:
                    continue

                result: dict[str, object] = {
                    "package": display_name,
                    "package_name": pkg_name,
                    "success": False,
                    "dry_run": dry_run,
                    "error": None,
                    "provisioned": True,
                }

                if dry_run:
                    _logger.info(
                        "[DRY-RUN] Would remove provisioned AppX package: %s",
                        display_name,
                    )
                    result["success"] = True
                    results.append(result)
                    continue

                try:
                    remove_cmd = f'Remove-AppxProvisionedPackage -Online -PackageName "{pkg_name}"'
                    _logger.info("Removing provisioned AppX package: %s", display_name)
                    remove_result = _run_powershell(remove_cmd, timeout=300)

                    if remove_result.returncode == 0:
                        result["success"] = True
                        _logger.info(
                            "Successfully removed provisioned AppX package: %s",
                            display_name,
                        )
                    else:
                        result["error"] = remove_result.stderr.strip() or "Unknown error"
                        _logger.warning(
                            "Failed to remove provisioned AppX package %s: %s",
                            display_name,
                            result["error"],
                        )

                except subprocess.TimeoutExpired:
                    result["error"] = "Timeout"
                    _logger.warning("Timeout removing provisioned AppX package: %s", display_name)
                except OSError as e:
                    result["error"] = str(e)
                    _logger.warning(
                        "Error removing provisioned AppX package %s: %s", display_name, e
                    )

                results.append(result)

        except subprocess.TimeoutExpired:
            _logger.warning("Timeout querying provisioned AppX packages for pattern: %s", pattern)
        except OSError as e:
            _logger.warning("Failed to query provisioned AppX packages: %s", e)

    return results


def is_office_store_install() -> bool:
    """!
    @brief Check if Office is installed via Microsoft Store.
    @returns True if any Office AppX packages are detected.
    """
    packages = detect_office_appx_packages()
    return len(packages) > 0
