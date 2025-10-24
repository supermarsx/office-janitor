"""!
@brief Detection helpers for installed Microsoft Office components.
@details The detection pipeline queries registry hives, filesystem locations,
and running processes to assemble an inventory of MSI and Click-to-Run Office
deployments as described in the project specification.
"""
from __future__ import annotations

from pathlib import Path
from typing import Dict, List

from . import constants, registry_tools


def _friendly_channel(raw_channel: str | None) -> str:
    """!
    @brief Resolve a friendly channel name for Click-to-Run metadata.
    """

    if not raw_channel:
        return "unknown"
    return constants.C2R_CHANNELS.get(raw_channel, raw_channel)


def detect_msi_installations() -> List[Dict[str, str]]:
    """!
    @brief Inspect the registry and return metadata for MSI-based Office installs.
    """

    installations: List[Dict[str, str]] = []
    seen_codes: set[str] = set()

    for root, base_key in constants.MSI_UNINSTALL_ROOTS:
        try:
            subkeys = list(registry_tools.iter_subkeys(root, base_key))
        except (FileNotFoundError, OSError):
            continue

        for subkey in subkeys:
            key_path = f"{base_key}\\{subkey}"
            entry = registry_tools.read_values(root, key_path)
            if not entry:
                continue
            product_code = entry.get("ProductCode") or subkey
            metadata = constants.MSI_PRODUCT_CODES.get(product_code)
            if not metadata or product_code in seen_codes:
                continue

            installation: Dict[str, str] = {
                "product_code": product_code,
                "release": metadata.get("release", "unknown"),
                "generation": metadata.get("generation", "unknown"),
                "edition": metadata.get("edition", entry.get("DisplayName", "")),
                "display_name": entry.get("DisplayName", metadata.get("edition", "")),
                "display_version": entry.get("DisplayVersion", ""),
                "architecture": metadata.get("architecture", "unknown"),
                "channel": "MSI",
                "uninstall_key": key_path,
                "source": f"{registry_tools.hive_name(root)}\\{key_path}",
            }
            install_root = metadata.get("install_path")
            if install_root:
                installation["install_path"] = install_root

            installations.append(installation)
            seen_codes.add(product_code)

    return installations


def detect_c2r_installations() -> List[Dict[str, object]]:
    """!
    @brief Probe Click-to-Run configuration to describe installed suites.
    """

    installations: List[Dict[str, object]] = []

    for root, config_path in constants.C2R_CONFIGURATION_KEYS:
        config_values = registry_tools.read_values(root, config_path)
        if not config_values:
            continue

        raw_release_ids = str(config_values.get("ProductReleaseIds", "")).split(",")
        release_ids = [rid.strip() for rid in raw_release_ids if rid.strip()]
        if not release_ids:
            continue

        platform = str(config_values.get("Platform") or config_values.get("PlatformId") or "").lower()
        architecture = constants.C2R_PLATFORM_ALIASES.get(platform, platform or "unknown")
        version = (
            config_values.get("VersionToReport")
            or config_values.get("ClientVersionToReport")
            or config_values.get("ProductVersion")
            or ""
        )
        channel_identifier = (
            config_values.get("UpdateChannel")
            or config_values.get("ChannelId")
            or config_values.get("CDNBaseUrl")
        )
        channel = _friendly_channel(str(channel_identifier) if channel_identifier else None)

        subscriptions: List[Dict[str, str]] = []
        for sub_root, sub_path in constants.C2R_SUBSCRIPTION_ROOTS:
            try:
                subkeys = list(registry_tools.iter_subkeys(sub_root, sub_path))
            except (FileNotFoundError, OSError):
                continue
            for subkey in subkeys:
                values = registry_tools.read_values(sub_root, f"{sub_path}\\{subkey}")
                raw_channel = values.get("ChannelId") or values.get("UpdateChannel")
                subscriptions.append(
                    {
                        "product_id": subkey,
                        "channel": _friendly_channel(str(raw_channel) if raw_channel else None),
                    }
                )

        com_entries: List[str] = []
        for com_root, com_path in constants.C2R_COM_REGISTRY_PATHS:
            try:
                com_entries.extend(list(registry_tools.iter_subkeys(com_root, com_path)))
            except (FileNotFoundError, OSError):
                continue

        installations.append(
            {
                "release_ids": release_ids,
                "channel": channel,
                "architecture": architecture,
                "version": str(version),
                "com_registration_count": len(com_entries),
                "subscriptions": subscriptions,
                "source": f"{registry_tools.hive_name(root)}\\{config_path}",
            }
        )

    return installations


def gather_office_inventory() -> Dict[str, List[Dict[str, object]]]:
    """!
    @brief Aggregate MSI, C2R, and ancillary signals into an inventory payload.
    """

    inventory: Dict[str, List[Dict[str, object]]] = {
        "msi": detect_msi_installations(),
        "c2r": detect_c2r_installations(),
        "filesystem": [],
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
