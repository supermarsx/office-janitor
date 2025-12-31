"""!
@brief Helper utilities shared by native OffScrub flows.
@details Encapsulates legacy argument parsing, execution directive derivation,
target selection, and optional cleanup steps so :mod:`off_scrub_native` can
focus on orchestration.
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, List, Mapping, MutableMapping, Sequence

from . import constants, detect, fs_tools, logging_ext, registry_tools, tasks_services

_GUID_PATTERN = re.compile(
    r"{?[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}}?"
)
_MSI_MAJOR_VERSION_MAP = {
    "11": "2003",
    "12": "2007",
    "14": "2010",
    "15": "2013",
    "16": "2016",
}
_SCRIPT_VERSION_HINTS = {
    "offscrub03.vbs": "2003",
    "offscrub07.vbs": "2007",
    "offscrub10.vbs": "2010",
    "offscrub_o15msi.vbs": "2013",
    "offscrub_o16msi.vbs": "2016",
    "offscrubc2r.vbs": "c2r",
}
_HIVE_NAMES = {value: key for key, value in constants.REGISTRY_ROOTS.items()}
_ADDIN_VERSION_KEYS = ("11.0", "12.0", "14.0", "15.0", "16.0")
_USER_SETTINGS_PATHS = (
    r"%APPDATA%\\Microsoft\\Office",
    r"%LOCALAPPDATA%\\Microsoft\\Office",
    r"%APPDATA%\\Microsoft\\Templates",
    r"%LOCALAPPDATA%\\Microsoft\\Office\\Templates",
)
_VBA_PATHS = (
    r"%APPDATA%\\Microsoft\\VBA",
    r"%LOCALAPPDATA%\\Microsoft\\VBA",
)
_SHORTCUT_PATHS = (
    r"%PROGRAMDATA%\\Microsoft\\Windows\\Start Menu\\Programs\\Microsoft Office",
    r"%PROGRAMDATA%\\Microsoft\\Windows\\Start Menu\\Programs\\Microsoft Office Tools",
    r"%APPDATA%\\Microsoft\\Windows\\Start Menu\\Programs\\Microsoft Office",
    r"%APPDATA%\\Microsoft\\Windows\\Start Menu\\Programs\\Microsoft Office Tools",
    r"%APPDATA%\\Microsoft\\Internet Explorer\\Quick Launch\\User Pinned\\TaskBar",
    r"%APPDATA%\\Microsoft\\Internet Explorer\\Quick Launch\\User Pinned\\StartMenu",
)


@dataclass
class LegacyInvocation:
    """!
    @brief Parsed legacy OffScrub invocation details.
    @details Captures the script path, implied version group, and recognised
    legacy flags so the native implementation can reproduce VBS semantics.
    """

    script_path: Path | None
    version_group: str | None
    product_codes: List[str]
    release_ids: List[str]
    flags: MutableMapping[str, object]
    unknown: List[str]
    log_directory: Path | None = None


@dataclass
class ExecutionDirectives:
    """!
    @brief Normalised behaviours derived from legacy flags.
    @details Records how many reruns to perform and which optional behaviours
    should be toggled for compatibility (e.g. skipping shortcut detection).
    """

    reruns: int = 1
    keep_license: bool = False
    skip_shortcut_detection: bool = False
    offline: bool = False
    quiet: bool = False
    no_reboot: bool = False
    delete_user_settings: bool = False
    keep_user_settings: bool = False
    clear_addin_registry: bool = False
    remove_vba: bool = False
    return_error_or_success: bool = False


def normalize_guid_token(token: str) -> str:
    """!
    @brief Normalise GUID-like strings into ``{GUID}`` form when possible.
    """

    cleaned = token.strip().strip("\0")
    if not cleaned:
        return ""
    trimmed = cleaned.strip("{}")
    candidate = f"{{{trimmed.upper()}}}" if trimmed else cleaned.upper()
    if _GUID_PATTERN.fullmatch(cleaned) or _GUID_PATTERN.fullmatch(candidate):
        return candidate
    return cleaned


def infer_version_group_from_script(script_path: Path | None, default: str | None = None) -> str | None:
    """!
    @brief Infer an OffScrub version group from the script filename.
    """

    if script_path is None:
        return default
    name = script_path.name.lower()
    return _SCRIPT_VERSION_HINTS.get(name, default)


def parse_legacy_arguments(command: str, argv: Sequence[str]) -> LegacyInvocation:
    """!
    @brief Parse legacy VBS-style arguments into a structured representation.
    """

    script_path: Path | None = None
    flags: MutableMapping[str, object] = {}
    unknown: List[str] = []
    product_codes: List[str] = []
    release_ids: List[str] = []
    log_directory: Path | None = None
    tokens = [str(part).strip() for part in argv if str(part).strip()]
    index = 0

    while index < len(tokens):
        raw_token = tokens[index]
        token = raw_token.strip()
        upper = token.upper()

        if script_path is None and upper.endswith(".VBS"):
            try:
                script_path = Path(token)
            except (TypeError, ValueError):
                script_path = None
            index += 1
            continue

        if token.startswith("/"):
            stripped = token.lstrip("/-").upper()
            if stripped == "ALL":
                flags["all"] = True
                index += 1
                continue
            if stripped == "OSE":
                flags["ose"] = True
                index += 1
                continue
            if stripped in _MSI_MAJOR_VERSION_MAP:
                flags["all"] = True
                flags["version_group"] = _MSI_MAJOR_VERSION_MAP[stripped]
                index += 1
                continue
            if stripped in ("OFFLINE", "FORCEOFFLINE"):
                flags["offline"] = True
                index += 1
                continue
            if stripped in ("QUIET", "PASSIVE"):
                flags["quiet"] = True
                index += 1
                continue
            if stripped in ("NOREBOOT", "NORESTART"):
                flags["no_reboot"] = True
                index += 1
                continue
            if stripped in ("PREVIEW", "DETECTONLY"):
                flags["detect_only"] = True
                index += 1
                continue
            if stripped in ("TR", "TESTRERUN"):
                flags["test_rerun"] = True
                index += 1
                continue
            if stripped in ("NOELEVATE", "NE"):
                flags["no_elevate"] = True
                index += 1
                continue
            if stripped in ("OFF", "OFFLINEONLY"):
                flags["offline"] = True
                index += 1
                continue
            if stripped in ("KL", "KEEPLICENSE"):
                flags["keep_license"] = True
                index += 1
                continue
            if stripped in ("RETERRORSUCCESS", "RETURNERRORORSUCCESS", "REOS"):
                flags["return_error_or_success"] = True
                index += 1
                continue
            if stripped in ("SKIPSD", "S", "SKIPSHORTCUTDETECTION"):
                flags["skip_shortcut_detection"] = True
                index += 1
                continue
            if stripped in ("TESTRERUN", "TR"):
                flags["test_rerun"] = True
                index += 1
                continue
            if stripped in ("BYPASS", "B"):
                flags["bypass"] = True
                index += 1
                continue
            if stripped in ("LOG", "L"):
                if index + 1 < len(tokens):
                    try:
                        log_directory = Path(tokens[index + 1].strip('"'))
                    except (TypeError, ValueError):
                        log_directory = None
                    index += 2
                    continue
            if stripped in ("DELETEUSERSETTINGS", "DUS"):
                flags["delete_user_settings"] = True
                index += 1
                continue
            if stripped in ("KEEPUSERSETTINGS", "KUS"):
                flags["keep_user_settings"] = True
                index += 1
                continue
            if stripped in ("CLEARADDINREG", "CAR"):
                flags["clear_addin_registry"] = True
                index += 1
                continue
            if stripped in ("REMOVEVBA",):
                flags["remove_vba"] = True
                index += 1
                continue
            if stripped in ("OSE", "O"):
                flags["ose"] = True
                index += 1
                continue
            if stripped in ("FORCE", "F"):
                flags["force"] = True
                index += 1
                continue
            if stripped in ("FASTREMOVE", "FR"):
                flags["fast_remove"] = True
                index += 1
                continue
            if stripped in ("SCANCOMPONENTS", "SC"):
                flags["scan_components"] = True
                index += 1
                continue
            unknown.append(token)
            index += 1
            continue

        if _GUID_PATTERN.fullmatch(token):
            normalized = normalize_guid_token(token)
            product_codes.append(normalized)
            index += 1
            continue

        if token.upper() == "ALL":
            flags["all"] = True
            index += 1
            continue

        if command == "c2r":
            release_ids.append(token)
        else:
            normalized = normalize_guid_token(token)
            product_codes.append(normalized)
        index += 1

    version_group = infer_version_group_from_script(
        script_path, default=("c2r" if command == "c2r" else None)
    )

    return LegacyInvocation(
        script_path=script_path,
        version_group=version_group,
        product_codes=product_codes,
        release_ids=release_ids,
        flags=flags,
        unknown=unknown,
        log_directory=log_directory,
    )


def derive_execution_directives(legacy: LegacyInvocation, *, dry_run: bool) -> ExecutionDirectives:
    """!
    @brief Translate legacy flags into execution directives for the native flow.
    """

    reruns = 2 if legacy.flags.get("test_rerun") else 1
    return ExecutionDirectives(
        reruns=reruns,
        keep_license=bool(legacy.flags.get("keep_license")),
        skip_shortcut_detection=bool(legacy.flags.get("skip_shortcut_detection")),
        offline=bool(legacy.flags.get("offline")),
        quiet=bool(legacy.flags.get("quiet")),
        no_reboot=bool(legacy.flags.get("no_reboot")),
        delete_user_settings=bool(legacy.flags.get("delete_user_settings") and not legacy.flags.get("keep_user_settings")),
        keep_user_settings=bool(legacy.flags.get("keep_user_settings")),
        clear_addin_registry=bool(legacy.flags.get("clear_addin_registry")),
        remove_vba=bool(legacy.flags.get("remove_vba")),
        return_error_or_success=bool(legacy.flags.get("return_error_or_success")),
    )


def _infer_c2r_group(entry: Mapping[str, object]) -> str | None:
    properties = entry.get("properties") if isinstance(entry, Mapping) else None
    if isinstance(properties, Mapping):
        supported_versions = properties.get("supported_versions")
        if isinstance(supported_versions, Sequence):
            for candidate in supported_versions:
                if str(candidate).strip():
                    return str(candidate).strip().lower()
    version = entry.get("version")
    if isinstance(version, str) and version.strip():
        major = version.strip().split(".", 1)[0]
        if major in _MSI_MAJOR_VERSION_MAP:
            return _MSI_MAJOR_VERSION_MAP[major]
    return None


def select_msi_targets(invocation: LegacyInvocation, inventory: Mapping[str, object]) -> List[Mapping[str, object]]:
    """!
    @brief Filter MSI inventory entries based on legacy invocation hints.
    """

    products: List[Mapping[str, object]] = []
    desired_codes = {code.upper() for code in invocation.product_codes}
    version_group = invocation.version_group
    msi_entries = inventory.get("msi", []) if isinstance(inventory, Mapping) else []

    for entry in msi_entries:
        if not isinstance(entry, Mapping):
            continue
        product_code = str(entry.get("product_code") or "").upper()
        if desired_codes and product_code not in desired_codes:
            continue
        properties = entry.get("properties", {})
        supported_versions = set(str(ver) for ver in properties.get("supported_versions", []) if str(ver))
        if version_group and supported_versions and version_group not in supported_versions:
            continue
        products.append(dict(entry))

    if not products and desired_codes:
        products = [{"product_code": code} for code in desired_codes]

    return products


def select_c2r_targets(invocation: LegacyInvocation, inventory: Mapping[str, object]) -> List[Mapping[str, object]]:
    """!
    @brief Filter Click-to-Run inventory entries based on legacy invocation hints.
    """

    targets: List[Mapping[str, object]] = []
    available = inventory.get("c2r", []) if isinstance(inventory, Mapping) else []
    desired_release_ids = {rid.lower() for rid in invocation.release_ids if rid}
    allow_all = bool(invocation.flags.get("all") or desired_release_ids)

    for entry in available:
        if not isinstance(entry, Mapping):
            continue
        releases = [
            str(rid)
            for rid in entry.get("release_ids", [])
            if str(rid).strip()
        ]
        if desired_release_ids and not any(rid.lower() in desired_release_ids for rid in releases):
            continue
        group = _infer_c2r_group(entry)
        if invocation.version_group and invocation.version_group != "c2r" and group and group != invocation.version_group:
            continue
        if not allow_all and invocation.version_group is None:
            continue

        targets.append(dict(entry))

    if not targets and desired_release_ids:
        targets.append({"release_ids": list(invocation.release_ids)})

    return targets


def format_registry_keys(entries: Iterable[object]) -> List[str]:
    """!
    @brief Convert registry entry tuples or strings into canonical ``HK**\\path`` text.
    """

    formatted: List[str] = []
    for entry in entries:
        if isinstance(entry, str):
            formatted.append(entry)
            continue
        if isinstance(entry, tuple) and len(entry) == 2:
            hive, path = entry
            name = _HIVE_NAMES.get(hive)
            if name:
                formatted.append(f"{name}\\{str(path).strip('\\\\')}")
    return formatted


def perform_optional_cleanup(directives: ExecutionDirectives, *, dry_run: bool, kind: str | None = None) -> None:
    """!
    @brief Execute optional cleanup implied by legacy flags.
    @param kind Optional legacy command identifier (``c2r`` or ``msi``) used to scope cleanup.
    """

    human_logger = logging_ext.get_human_logger()

    if kind == "c2r":
        human_logger.info("Removing Click-to-Run scheduled tasks referenced by legacy scripts.")
        tasks_services.delete_tasks(constants.C2R_CLEANUP_TASKS, dry_run=dry_run)
        human_logger.info("Cleaning Click-to-Run COM compatibility registry keys.")
        com_keys = format_registry_keys(constants.C2R_COM_REGISTRY_PATHS)
        try:
            registry_tools.delete_keys(com_keys, dry_run=dry_run)
        except registry_tools.RegistryError as exc:  # pragma: no cover - defensive
            human_logger.warning("Click-to-Run COM registry cleanup skipped: %s", exc)
        if directives.keep_license:
            human_logger.info("Skipping Click-to-Run cache cleanup because keep-license was requested.")
        else:
            c2r_residue = [
                str(Path(os.path.expandvars(entry["path"])))
                for entry in constants.RESIDUE_PATH_TEMPLATES
                if entry.get("category") == "c2r_cache"
            ]
            if c2r_residue:
                human_logger.info("Removing Click-to-Run cache directories referenced by legacy scripts.")
                fs_tools.remove_paths(c2r_residue, dry_run=dry_run)

    if not directives.skip_shortcut_detection:
        human_logger.info("Removing legacy Office shortcuts from known Start Menu roots.")
        fs_tools.remove_paths(_SHORTCUT_PATHS, dry_run=dry_run)
    else:
        human_logger.info("Skipping shortcut cleanup per legacy SkipSD flag.")

    residue_keys = format_registry_keys(constants.REGISTRY_RESIDUE_PATHS)
    if residue_keys:
        human_logger.info("Removing legacy Office registry residue keys.")
        try:
            registry_tools.delete_keys(residue_keys, dry_run=dry_run)
        except registry_tools.RegistryError as exc:  # pragma: no cover - defensive
            human_logger.warning("Registry residue cleanup skipped: %s", exc)

    if directives.delete_user_settings and not directives.keep_user_settings:
        human_logger.info("Deleting user settings directories requested by legacy flags.")
        fs_tools.remove_paths(_USER_SETTINGS_PATHS, dry_run=dry_run)

    if directives.clear_addin_registry:
        human_logger.info("Clearing Office add-in registry keys requested by legacy flags.")
        addin_keys = []
        for version in _ADDIN_VERSION_KEYS:
            addin_keys.extend(
                [
                    f"HKCU\\Software\\Microsoft\\Office\\{version}\\Addins",
                    f"HKLM\\SOFTWARE\\Microsoft\\Office\\{version}\\Addins",
                    f"HKLM\\SOFTWARE\\WOW6432Node\\Microsoft\\Office\\{version}\\Addins",
                ]
            )
        try:
            registry_tools.delete_keys(addin_keys, dry_run=dry_run)
        except registry_tools.RegistryError as exc:  # pragma: no cover - defensive
            human_logger.warning("Add-in registry cleanup skipped: %s", exc)

    if directives.remove_vba:
        human_logger.info("Removing VBA registry keys requested by legacy flags.")
        vba_keys = []
        for version in _ADDIN_VERSION_KEYS:
            vba_keys.extend(
                [
                    f"HKCU\\Software\\Microsoft\\Office\\{version}\\VBA",
                    f"HKLM\\SOFTWARE\\Microsoft\\Office\\{version}\\VBA",
                    f"HKLM\\SOFTWARE\\WOW6432Node\\Microsoft\\Office\\{version}\\VBA",
                ]
            )
        try:
            registry_tools.delete_keys(vba_keys, dry_run=dry_run)
        except registry_tools.RegistryError as exc:  # pragma: no cover - defensive
            human_logger.warning("VBA registry cleanup skipped: %s", exc)
        human_logger.info("Removing VBA filesystem caches requested by legacy flags.")
        fs_tools.remove_paths(_VBA_PATHS, dry_run=dry_run)
