"""!
@brief Detection helpers for MSI and Click-to-Run Office deployments.
@details Reads structured metadata from :mod:`office_janitor.constants`, probes
registry hives, and returns structured :class:`DetectedInstallation` records that
contain uninstall handles, source type, and channel information.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Mapping, Tuple

from . import constants, registry_tools


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

            properties: Dict[str, object] = {
                "display_name": display_name,
                "display_version": display_version,
            }
            if uninstall_string:
                properties["uninstall_string"] = uninstall_string
            properties["supported_versions"] = list(metadata.get("supported_versions", ()))
            properties["edition"] = metadata.get("edition", "")

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


def gather_office_inventory() -> Dict[str, List[Dict[str, object]]]:
    """!
    @brief Aggregate MSI, C2R, and ancillary signals into an inventory payload.
    """

    inventory: Dict[str, List[Dict[str, object]]] = {
        "msi": [entry.to_dict() for entry in detect_msi_installations()],
        "c2r": [entry.to_dict() for entry in detect_c2r_installations()],
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


def reprobe(options: Mapping[str, object] | None = None) -> Dict[str, List[Dict[str, object]]]:
    """!
    @brief Re-run Office detection after a scrub pass to check for leftovers.
    @details The optional ``options`` mapping is accepted for parity with future
    targeted detection strategies, but currently serves only as a hook for
    logging and diagnostics. The returned inventory mirrors
    :func:`gather_office_inventory`.
    """

    _ = options  # Options are presently unused but reserved for parity.
    return gather_office_inventory()
