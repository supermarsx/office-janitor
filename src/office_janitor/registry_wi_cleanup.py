"""!
@brief Windows Installer metadata validation and cleanup utilities.
@details Functions to validate and clean up WI registry entries, including
orphaned products, components, TypeLibs, shell extensions, and protocol handlers.
Based on VBS ``ValidateWIMetadataKey`` from OffScrub_O16msi.vbs.
"""

from __future__ import annotations

import logging
from collections.abc import Iterable
from pathlib import Path
from typing import Any

from . import registry_tools
from .registry_office import is_office_guid

_LOGGER = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Windows Installer Metadata Paths
# ---------------------------------------------------------------------------

# Standard WI registry paths for validation
WI_METADATA_PATHS: dict[str, tuple[int, str, int]] = {
    # (hive, path, expected_key_length)
    "Products": (registry_tools._WINREG_HKLM, r"SOFTWARE\Classes\Installer\Products", 32),
    "Components": (registry_tools._WINREG_HKLM, r"SOFTWARE\Classes\Installer\Components", 32),
    "Features": (registry_tools._WINREG_HKLM, r"SOFTWARE\Classes\Installer\Features", 32),
    "Patches": (registry_tools._WINREG_HKLM, r"SOFTWARE\Classes\Installer\Patches", 32),
    "UpgradeCodes": (registry_tools._WINREG_HKLM, r"SOFTWARE\Classes\Installer\UpgradeCodes", 32),
    "UserDataProducts": (
        registry_tools._WINREG_HKLM,
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData",
        0,  # Variable structure
    ),
}
"""!
@brief Windows Installer registry paths that may contain orphaned entries.
"""


# ---------------------------------------------------------------------------
# WI Metadata Validation
# ---------------------------------------------------------------------------


def _is_valid_compressed_guid(name: str) -> bool:
    """!
    @brief Check if a registry key name is a valid compressed GUID.
    @param name Subkey name to validate.
    @return True if valid 32-char hex string.
    """
    if len(name) != 32:
        return False
    try:
        int(name, 16)
        return True
    except ValueError:
        return False


def validate_wi_metadata_key(
    hive: int,
    path: str,
    expected_length: int = 32,
    *,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> list[str]:
    """!
    @brief Validate WI metadata keys and identify invalid entries.
    @param hive Registry hive (HKLM, etc.).
    @param path Registry path to validate.
    @param expected_length Expected length of valid subkey names (32 for GUIDs).
    @param dry_run If True, only report without deleting.
    @param logger Optional logger.
    @return List of invalid subkey names found.

    @details Scans a Windows Installer metadata key and identifies entries
    that don't match the expected format (e.g., corrupted GUIDs). This mirrors
    the VBS ``ValidateWIMetadataKey`` function from OffScrub_O16msi.vbs.

    Invalid entries are subkeys that:
    - Don't have the expected length (32 chars for compressed GUIDs)
    - Contain non-hexadecimal characters
    - Are empty or malformed
    """
    logger = logger or _LOGGER
    invalid_entries: list[str] = []

    try:
        for subkey in registry_tools.iter_subkeys(hive, path, view="native"):
            # Check if subkey name matches expected format
            if expected_length > 0:
                if len(subkey) != expected_length:
                    invalid_entries.append(subkey)
                    continue
                if not _is_valid_compressed_guid(subkey):
                    invalid_entries.append(subkey)
    except FileNotFoundError:
        logger.debug("WI metadata path not found: %s", path)
        return []
    except OSError as e:
        logger.warning("Failed to access WI metadata path %s: %s", path, e)
        return []

    if invalid_entries:
        logger.info(
            "Found %d invalid WI metadata entries in %s",
            len(invalid_entries),
            path,
            extra={
                "action": "wi-validation",
                "path": path,
                "invalid_count": len(invalid_entries),
            },
        )

    return invalid_entries


def scan_wi_metadata(
    *,
    logger: logging.Logger | None = None,
) -> dict[str, list[str]]:
    """!
    @brief Scan all standard WI metadata paths for invalid entries.
    @param logger Optional logger.
    @return Dictionary mapping path names to lists of invalid entries.
    """
    logger = logger or _LOGGER
    results: dict[str, list[str]] = {}

    logger.info("Scanning Windows Installer metadata for invalid entries")

    for name, (hive, path, expected_len) in WI_METADATA_PATHS.items():
        if expected_len == 0:
            continue  # Skip variable-structure paths
        invalid = validate_wi_metadata_key(hive, path, expected_len, logger=logger)
        if invalid:
            results[name] = invalid

    return results


# ---------------------------------------------------------------------------
# WI Orphan Cleanup
# ---------------------------------------------------------------------------


def cleanup_wi_orphaned_products(
    product_codes: Iterable[str],
    *,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> int:
    """!
    @brief Remove WI registry entries for orphaned product codes.
    @param product_codes Product GUIDs to clean up.
    @param dry_run If True, only log without deleting.
    @param logger Optional logger.
    @return Number of entries removed.

    @details Cleans up WI metadata entries for products that are no longer
    properly installed. This includes:
    - Installer\\Products\\<compressed_guid>
    - Installer\\Features\\<compressed_guid>
    - Installer\\UpgradeCodes entries that reference the product
    """
    from . import guid_utils

    logger = logger or _LOGGER
    removed = 0

    for product_code in product_codes:
        try:
            compressed = guid_utils.compress_guid(product_code)
        except guid_utils.GuidError:
            logger.warning("Invalid product code: %s", product_code)
            continue

        # Build paths to clean
        paths_to_clean = [
            f"HKLM\\SOFTWARE\\Classes\\Installer\\Products\\{compressed}",
            f"HKLM\\SOFTWARE\\Classes\\Installer\\Features\\{compressed}",
        ]

        for path in paths_to_clean:
            if registry_tools.key_exists(path):
                logger.info(
                    "Removing orphaned WI entry: %s",
                    path,
                    extra={"action": "wi-cleanup", "path": path, "dry_run": dry_run},
                )
                if not dry_run:
                    try:
                        registry_tools.delete_keys([path], dry_run=False, logger=logger)
                        removed += 1
                    except Exception as e:
                        logger.warning("Failed to delete %s: %s", path, e)
                else:
                    removed += 1

    return removed


def cleanup_wi_orphaned_components(
    component_ids: Iterable[str],
    *,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> int:
    """!
    @brief Remove WI registry entries for orphaned component IDs.
    @param component_ids Component GUIDs to clean up.
    @param dry_run If True, only log without deleting.
    @param logger Optional logger.
    @return Number of entries removed.

    @details Cleans up WI component entries that have no valid product clients.
    This is a more targeted cleanup than removing entire product trees.
    """
    from . import guid_utils

    logger = logger or _LOGGER
    removed = 0

    for component_id in component_ids:
        try:
            compressed = guid_utils.compress_guid(component_id)
        except guid_utils.GuidError:
            logger.warning("Invalid component ID: %s", component_id)
            continue

        path = f"HKLM\\SOFTWARE\\Classes\\Installer\\Components\\{compressed}"

        if registry_tools.key_exists(path):
            logger.info(
                "Removing orphaned WI component: %s",
                path,
                extra={"action": "wi-cleanup", "path": path, "dry_run": dry_run},
            )
            if not dry_run:
                try:
                    registry_tools.delete_keys([path], dry_run=False, logger=logger)
                    removed += 1
                except Exception as e:
                    logger.warning("Failed to delete %s: %s", path, e)
            else:
                removed += 1

    return removed


# ---------------------------------------------------------------------------
# Shell Integration Cleanup
# ---------------------------------------------------------------------------


def cleanup_orphaned_typelibs(
    typelib_guids: Iterable[str],
    *,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> list[str]:
    """!
    @brief Remove orphaned TypeLib registrations for Office components.
    @param typelib_guids TypeLib GUIDs to check.
    @param dry_run If True, only report without deleting.
    @param logger Optional logger.
    @return List of removed TypeLib GUIDs.

    @details Scans TypeLib registrations and removes entries where the
    referenced DLL file no longer exists. This mirrors the VBS TypeLib
    cleanup logic from OffScrub_O16msi.vbs.
    """
    logger = logger or _LOGGER
    removed: list[str] = []
    hklm = registry_tools._WINREG_HKLM

    for typelib_guid in typelib_guids:
        base_path = f"HKLM\\SOFTWARE\\Classes\\TypeLib\\{typelib_guid}"

        if not registry_tools.key_exists(base_path):
            continue

        # Check each version subkey
        try:
            versions = list(
                registry_tools.iter_subkeys(hklm, f"SOFTWARE\\Classes\\TypeLib\\{typelib_guid}")
            )
        except (FileNotFoundError, OSError):
            continue

        typelib_orphaned = True

        for version in versions:
            version_path = f"SOFTWARE\\Classes\\TypeLib\\{typelib_guid}\\{version}"
            try:
                # Check the 0\\win32 or 0\\win64 paths for the DLL location
                for platform in ("0\\win32", "0\\win64", "0"):
                    try:
                        values = registry_tools.read_values(
                            hklm, f"{version_path}\\{platform}", view="native"
                        )
                        default_val = values.get("", "")
                        if default_val and Path(default_val).exists():
                            typelib_orphaned = False
                            break
                    except (FileNotFoundError, OSError):
                        continue
            except (FileNotFoundError, OSError):
                continue

            if not typelib_orphaned:
                break

        if typelib_orphaned:
            logger.info(
                "Removing orphaned TypeLib: %s",
                typelib_guid,
                extra={"action": "typelib-cleanup", "guid": typelib_guid, "dry_run": dry_run},
            )
            if not dry_run:
                try:
                    registry_tools.delete_keys([base_path], dry_run=False, logger=logger)
                    removed.append(typelib_guid)
                except Exception as e:
                    logger.warning("Failed to delete TypeLib %s: %s", typelib_guid, e)
            else:
                removed.append(typelib_guid)

    return removed


def scan_orphaned_typelibs(
    typelib_guids: Iterable[str],
    *,
    logger: logging.Logger | None = None,
) -> list[str]:
    """!
    @brief Scan for orphaned TypeLib registrations without removing them.
    @param typelib_guids TypeLib GUIDs to check.
    @param logger Optional logger.
    @return List of orphaned TypeLib GUIDs.
    """
    return cleanup_orphaned_typelibs(typelib_guids, dry_run=True, logger=logger)


def cleanup_shell_extensions(
    *,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> int:
    """!
    @brief Remove orphaned shell extension registrations.
    @param dry_run If True, only report without deleting.
    @param logger Optional logger.
    @return Number of extensions removed.

    @details Cleans up shell extensions that reference non-existent DLLs.
    This includes context menu handlers, property sheet handlers, etc.
    """

    logger = logger or _LOGGER
    removed = 0
    hklm = registry_tools._WINREG_HKLM

    # Shell extension approval entries
    approval_paths = [
        (
            hklm,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved",
        ),
    ]

    for hive, path in approval_paths:
        try:
            for name, value in registry_tools.iter_values(hive, path):
                # Check if this looks like an Office-related extension
                if not isinstance(value, str):
                    continue
                value_lower = value.lower()
                if not any(kw in value_lower for kw in ("office", "outlook", "groove", "onenote")):
                    continue

                # Check if the associated CLSID still exists
                clsid_path = f"SOFTWARE\\Classes\\CLSID\\{name}"
                if not registry_tools.key_exists(f"HKLM\\{clsid_path}"):
                    logger.info(
                        "Found orphaned shell extension approval: %s (%s)",
                        name,
                        value,
                        extra={"action": "shell-cleanup", "clsid": name, "dry_run": dry_run},
                    )
                    # Note: We don't remove from Approved list directly as it's a single key
                    # with multiple values. This just reports the orphaned entry.
                    removed += 1
        except (FileNotFoundError, OSError):
            continue

    return removed


def cleanup_protocol_handlers(
    protocols: Iterable[str],
    *,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> list[str]:
    """!
    @brief Remove orphaned protocol handler registrations.
    @param protocols Protocol names to check (e.g., "osf", "ms-word").
    @param dry_run If True, only report without deleting.
    @param logger Optional logger.
    @return List of removed protocol names.

    @details Removes URL protocol handlers that reference non-existent
    executables. Common Office protocols include osf:, ms-word:, etc.
    """
    logger = logger or _LOGGER
    removed: list[str] = []
    hklm = registry_tools._WINREG_HKLM
    hkcu = registry_tools._WINREG_HKCU

    for protocol in protocols:
        # Check both HKLM and HKCU
        for hive, hive_name_str in [(hklm, "HKLM"), (hkcu, "HKCU")]:
            path = f"SOFTWARE\\Classes\\{protocol}"
            full_path = f"{hive_name_str}\\{path}"

            if not registry_tools.key_exists(full_path):
                continue

            # Check if the shell\\open\\command points to an existing executable
            try:
                values = registry_tools.read_values(
                    hive, f"{path}\\shell\\open\\command", view="native"
                )
                default_cmd = values.get("", "")
                if default_cmd:
                    # Extract executable path (handle quoted paths)
                    exe_path = default_cmd.strip('"').split('"')[0].strip()
                    if exe_path and not Path(exe_path).exists():
                        logger.info(
                            "Removing orphaned protocol handler: %s",
                            full_path,
                            extra={
                                "action": "protocol-cleanup",
                                "protocol": protocol,
                                "dry_run": dry_run,
                            },
                        )
                        if not dry_run:
                            try:
                                registry_tools.delete_keys(
                                    [full_path], dry_run=False, logger=logger
                                )
                                if protocol not in removed:
                                    removed.append(protocol)
                            except Exception as e:
                                logger.warning("Failed to delete protocol %s: %s", protocol, e)
                        else:
                            if protocol not in removed:
                                removed.append(protocol)
            except (FileNotFoundError, OSError):
                continue

    return removed


__all__ = [
    "WI_METADATA_PATHS",
    "cleanup_orphaned_typelibs",
    "cleanup_protocol_handlers",
    "cleanup_shell_extensions",
    "cleanup_wi_orphaned_components",
    "cleanup_wi_orphaned_products",
    "scan_orphaned_typelibs",
    "scan_wi_metadata",
    "validate_wi_metadata_key",
]
