"""!
@brief User registry hive management and vNext identity cleanup.
@details Functions for loading/unloading user registry hives and cleaning up
vNext identity, licensing, and taskband registry entries. Based on
OffScrubC2R.vbs LoadUsersReg and OfficeScrubber.cmd :vNextREG subroutines.
"""

from __future__ import annotations

import logging
import re
import shutil
from pathlib import Path
from typing import Any

from . import exec_utils, registry_tools

_LOGGER = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# vNext Identity Registry Cleanup Constants
# ---------------------------------------------------------------------------

_VNEXT_IDENTITY_KEYS: list[str] = [
    r"HKLM\SOFTWARE\Microsoft\Office\16.0\Common\Licensing",
    r"HKLM\SOFTWARE\Microsoft\Office\16.0\Common\Identity",
    r"HKLM\SOFTWARE\Microsoft\Office\16.0\Registration",
    r"HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Updates",
    r"HKLM\SOFTWARE\Microsoft\Office\16.0\Common\OEM",
    r"HKLM\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Licensing",
    # WOW6432Node variants
    r"HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Common\OEM",
    r"HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Common\Licensing",
    r"HKLM\SOFTWARE\Policies\WOW6432Node\Microsoft\Office\16.0\Common\Licensing",
]
"""!
@brief Registry keys to delete for vNext identity cleanup.
"""

_VNEXT_C2R_VALUES_TO_DELETE: list[str] = [
    "SharedComputerLicensing",
    "productkeys",
]
"""!
@brief Registry values to delete from ClickToRun Configuration key.
"""

_VNEXT_IDENTITY_VALUE_PATTERNS: list[str] = [
    r".*\.EmailAddress$",
    r".*\.TenantId$",
    r".*\.DeviceBasedLicensing$",
]
"""!
@brief Regex patterns for identity-related values to delete from C2R Configuration.
"""


# ---------------------------------------------------------------------------
# Registry Value Deletion
# ---------------------------------------------------------------------------


def delete_registry_value(
    key_path: str,
    value_name: str,
    *,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> bool:
    """!
    @brief Delete a specific registry value.
    @param key_path Full registry path (e.g., "HKLM\\SOFTWARE\\...").
    @param value_name Name of the value to delete.
    @param dry_run If True, only log without deleting.
    @param logger Optional logger.
    @returns True if deleted or dry-run, False if not found or error.
    """
    logger = logger or _LOGGER
    reg_executable = shutil.which("reg")

    if not reg_executable:
        logger.warning("reg.exe not found, cannot delete value")
        return False

    logger.info(
        "Deleting registry value",
        extra={
            "action": "registry-value-delete",
            "key": key_path,
            "value": value_name,
            "dry_run": dry_run,
        },
    )

    if dry_run:
        return True

    result = exec_utils.run_command(
        [reg_executable, "delete", key_path, "/v", value_name, "/f"],
        event="registry_value_delete",
        dry_run=False,
        check=False,
        extra={"key": key_path, "value": value_name},
    )

    return result.returncode == 0


# ---------------------------------------------------------------------------
# vNext Identity Cleanup
# ---------------------------------------------------------------------------


def cleanup_vnext_identity_registry(
    *,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> dict[str, Any]:
    """!
    @brief Clean up vNext identity and licensing registry entries.
    @details Implements OfficeScrubber.cmd :vNextREG subroutine functionality:
        1. Delete identity/licensing registry keys
        2. Delete specific C2R configuration values
        3. Delete identity-related values matching patterns (*.EmailAddress, etc.)
    @param dry_run If True, only log what would be deleted.
    @param logger Optional logger.
    @returns Dictionary with cleanup results.
    """
    logger = logger or _LOGGER
    results: dict[str, Any] = {
        "keys_deleted": [],
        "values_deleted": [],
        "patterns_matched": [],
        "errors": [],
    }

    # Step 1: Delete vNext identity keys
    logger.info("Cleaning vNext identity registry keys...")
    for key_path in _VNEXT_IDENTITY_KEYS:
        try:
            if registry_tools.key_exists(key_path):
                logger.debug("Deleting vNext key: %s", key_path)
                registry_tools.delete_keys([key_path], dry_run=dry_run, logger=logger)
                results["keys_deleted"].append(key_path)
        except Exception as e:
            logger.warning("Failed to delete key %s: %s", key_path, e)
            results["errors"].append({"key": key_path, "error": str(e)})

    # Step 2: Delete specific C2R configuration values
    c2r_config_paths = [
        r"HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
        r"HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration",
    ]

    for config_path in c2r_config_paths:
        if not registry_tools.key_exists(config_path):
            continue

        for value_name in _VNEXT_C2R_VALUES_TO_DELETE:
            try:
                if delete_registry_value(config_path, value_name, dry_run=dry_run, logger=logger):
                    results["values_deleted"].append(f"{config_path}\\{value_name}")
            except Exception as e:
                logger.debug("Failed to delete value %s\\%s: %s", config_path, value_name, e)

        # Step 3: Delete pattern-matched identity values
        # Need to enumerate values and match against patterns
        try:
            hive, _, subpath = config_path.partition("\\")
            hive_int = (
                registry_tools._WINREG_HKLM if hive == "HKLM" else registry_tools._WINREG_HKCU
            )
            patterns = [re.compile(p, re.IGNORECASE) for p in _VNEXT_IDENTITY_VALUE_PATTERNS]

            for value_name, _ in registry_tools.iter_values(hive_int, subpath):
                for pattern in patterns:
                    if pattern.match(value_name):
                        logger.debug("Pattern match for deletion: %s\\%s", config_path, value_name)
                        if delete_registry_value(
                            config_path, value_name, dry_run=dry_run, logger=logger
                        ):
                            results["patterns_matched"].append(f"{config_path}\\{value_name}")
                        break
        except (FileNotFoundError, OSError) as e:
            logger.debug("Failed to enumerate values in %s: %s", config_path, e)

    # Step 4: Clean SPP policies in Network Service SID (S-1-5-20)
    spp_key = (
        r"HKU\S-1-5-20\Software\Microsoft\OfficeSoftwareProtectionPlatform"
        r"\Policies\0ff1ce15-a989-479d-af46-f275c6370663"
    )
    if registry_tools.key_exists(spp_key):
        try:
            registry_tools.delete_keys([spp_key], dry_run=dry_run, logger=logger)
            results["keys_deleted"].append(spp_key)
        except Exception as e:
            logger.debug("Failed to delete SPP key %s: %s", spp_key, e)

    logger.info(
        "vNext identity cleanup complete: %d keys, %d values, %d pattern matches",
        len(results["keys_deleted"]),
        len(results["values_deleted"]),
        len(results["patterns_matched"]),
    )

    return results


# ---------------------------------------------------------------------------
# User Profile Registry Loading
# ---------------------------------------------------------------------------


def get_user_profiles_directory() -> Path | None:
    """!
    @brief Get the path to the Windows user profiles directory.
    @returns Path to profiles directory (e.g., C:\\Users) or None if not found.
    """
    try:
        import os

        key_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
        values = registry_tools.read_values(registry_tools._WINREG_HKLM, key_path, view="native")
        profiles_dir = values.get("ProfilesDirectory", "")
        if profiles_dir:
            # Expand environment variables
            expanded = os.path.expandvars(profiles_dir)
            return Path(expanded)
    except (FileNotFoundError, OSError):
        pass
    return None


def get_user_profile_hive_paths() -> list[tuple[str, Path]]:
    """!
    @brief Enumerate user profile folders and their ntuser.dat paths.
    @returns List of (profile_name, ntuser_dat_path) tuples.
    """
    profiles_dir = get_user_profiles_directory()
    if profiles_dir is None or not profiles_dir.exists():
        return []

    results: list[tuple[str, Path]] = []
    try:
        for folder in profiles_dir.iterdir():
            if not folder.is_dir():
                continue
            ntuser = folder / "ntuser.dat"
            if ntuser.exists():
                results.append((folder.name, ntuser))
    except OSError:
        pass

    return results


# Track loaded user hives for cleanup
_LOADED_USER_HIVES: list[str] = []


def load_user_registry_hives(
    *,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> list[str]:
    """!
    @brief Load all user ntuser.dat files into HKU for per-user cleanup.
    @details Implements OffScrubC2R.vbs LoadUsersReg functionality.
        Loads each user's registry hive to HKU\\<profile_name>.
    @param dry_run If True, only log what would be loaded.
    @param logger Optional logger.
    @returns List of successfully loaded hive names.
    """
    global _LOADED_USER_HIVES
    logger = logger or _LOGGER

    profiles = get_user_profile_hive_paths()
    if not profiles:
        logger.debug("No user profile hives found to load")
        return []

    loaded: list[str] = []
    reg_exe = shutil.which("reg")
    if not reg_exe:
        logger.warning("reg.exe not found, cannot load user hives")
        return []

    for profile_name, ntuser_path in profiles:
        hive_key = f"HKU\\{profile_name}"

        # Skip if already loaded (e.g., current user)
        if registry_tools.key_exists(hive_key):
            logger.debug("Hive %s already loaded, skipping", hive_key)
            continue

        logger.info(
            "Loading user registry hive",
            extra={
                "action": "registry-hive-load",
                "profile": profile_name,
                "path": str(ntuser_path),
                "dry_run": dry_run,
            },
        )

        if dry_run:
            loaded.append(profile_name)
            continue

        # reg load "HKU\<profile_name>" "<path>\ntuser.dat"
        result = exec_utils.run_command(
            [reg_exe, "load", hive_key, str(ntuser_path)],
            event="registry_hive_load",
            dry_run=False,
            check=False,
            extra={"profile": profile_name},
        )

        if result.returncode == 0:
            loaded.append(profile_name)
            _LOADED_USER_HIVES.append(profile_name)
            logger.debug("Loaded hive %s", hive_key)
        else:
            logger.debug(
                "Failed to load hive %s (code %d): %s",
                hive_key,
                result.returncode,
                result.stderr.strip() if result.stderr else "",
            )

    return loaded


def unload_user_registry_hives(
    *,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> list[str]:
    """!
    @brief Unload previously loaded user registry hives.
    @param dry_run If True, only log what would be unloaded.
    @param logger Optional logger.
    @returns List of successfully unloaded hive names.
    """
    global _LOADED_USER_HIVES
    logger = logger or _LOGGER

    if not _LOADED_USER_HIVES:
        logger.debug("No user hives to unload")
        return []

    unloaded: list[str] = []
    reg_exe = shutil.which("reg")
    if not reg_exe:
        logger.warning("reg.exe not found, cannot unload user hives")
        return []

    # Unload in reverse order
    for profile_name in reversed(_LOADED_USER_HIVES.copy()):
        hive_key = f"HKU\\{profile_name}"

        logger.info(
            "Unloading user registry hive",
            extra={
                "action": "registry-hive-unload",
                "profile": profile_name,
                "dry_run": dry_run,
            },
        )

        if dry_run:
            unloaded.append(profile_name)
            continue

        # reg unload "HKU\<profile_name>"
        result = exec_utils.run_command(
            [reg_exe, "unload", hive_key],
            event="registry_hive_unload",
            dry_run=False,
            check=False,
            extra={"profile": profile_name},
        )

        if result.returncode == 0:
            unloaded.append(profile_name)
            _LOADED_USER_HIVES.remove(profile_name)
            logger.debug("Unloaded hive %s", hive_key)
        else:
            logger.warning(
                "Failed to unload hive %s (code %d)",
                hive_key,
                result.returncode,
            )

    return unloaded


def get_loaded_user_hives() -> list[str]:
    """!
    @brief Get list of user hives currently loaded by this session.
    @returns List of profile names with loaded hives.
    """
    return list(_LOADED_USER_HIVES)


# ---------------------------------------------------------------------------
# Taskband Registry Cleanup
# ---------------------------------------------------------------------------

_TASKBAND_VALUES_TO_DELETE: list[str] = [
    "Favorites",
    "FavoritesRemovedChanges",
    "FavoritesChanges",
    "FavoritesResolve",
    "FavoritesVersion",
]
"""!
@brief Registry values to delete from Taskband key to unpin items.
"""


def cleanup_taskband_registry(
    *,
    include_all_users: bool = False,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> dict[str, Any]:
    """!
    @brief Clean up taskband registry to remove pinned Office items.
    @details Implements OffScrubC2R.vbs ClearTaskBand functionality.
        Removes Favorites* values from Taskband key to clear pinned items.
    @param include_all_users If True, also clean all user profiles in HKU.
    @param dry_run If True, only log what would be deleted.
    @param logger Optional logger.
    @returns Dictionary with cleanup results.
    """
    logger = logger or _LOGGER
    results: dict[str, Any] = {
        "values_deleted": [],
        "users_processed": [],
        "errors": [],
    }

    taskband_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband"
    hku = registry_tools._WINREG_HKU

    # Step 1: Clean HKCU taskband
    hkcu_full_path = f"HKCU\\{taskband_path}"
    logger.info("Cleaning HKCU taskband registry...")

    for value_name in _TASKBAND_VALUES_TO_DELETE:
        try:
            if delete_registry_value(hkcu_full_path, value_name, dry_run=dry_run, logger=logger):
                results["values_deleted"].append(f"{hkcu_full_path}\\{value_name}")
        except Exception as e:
            logger.debug("Failed to delete %s\\%s: %s", hkcu_full_path, value_name, e)

    # Step 2: If requested, clean all user profiles in HKU
    if include_all_users:
        logger.info("Cleaning taskband for all user profiles...")

        # First load user hives if not already loaded
        loaded_hives = load_user_registry_hives(dry_run=dry_run, logger=logger)
        if loaded_hives:
            results["users_processed"].extend(loaded_hives)

        # Enumerate all SIDs in HKU
        try:
            for sid in registry_tools.iter_subkeys(hku, "", view="native"):
                # Skip well-known SIDs that don't have user profiles
                if sid in ("S-1-5-18", "S-1-5-19", "S-1-5-20", ".DEFAULT"):
                    continue
                if sid.endswith("_Classes"):
                    continue

                hku_taskband_path = f"HKU\\{sid}\\{taskband_path}"

                for value_name in _TASKBAND_VALUES_TO_DELETE:
                    try:
                        if delete_registry_value(
                            hku_taskband_path, value_name, dry_run=dry_run, logger=logger
                        ):
                            results["values_deleted"].append(f"{hku_taskband_path}\\{value_name}")
                    except Exception:
                        pass  # Value may not exist, which is fine

                if sid not in results["users_processed"]:
                    results["users_processed"].append(sid)

        except (FileNotFoundError, OSError) as e:
            logger.debug("Failed to enumerate HKU: %s", e)

        # Unload hives we loaded
        if loaded_hives:
            unload_user_registry_hives(dry_run=dry_run, logger=logger)

    logger.info(
        "Taskband cleanup complete: %d values deleted across %d users",
        len(results["values_deleted"]),
        len(results["users_processed"]) + 1,  # +1 for HKCU
    )

    return results


__all__ = [
    "cleanup_taskband_registry",
    "cleanup_vnext_identity_registry",
    "delete_registry_value",
    "get_loaded_user_hives",
    "get_user_profile_hive_paths",
    "get_user_profiles_directory",
    "load_user_registry_hives",
    "unload_user_registry_hives",
]
