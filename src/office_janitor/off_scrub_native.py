"""!
@brief Native replacements for OffScrub helper flows (full parity).
@details Provides Python implementations that mirror the historical OffScrub
VBScript entrypoints, including legacy command-line flag parsing so existing
automation can migrate without ``cscript.exe``. The module can be invoked via
``python -m office_janitor.off_scrub_native`` or used programmatically.
"""
from __future__ import annotations

import argparse
import logging
import os
import sys
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, MutableMapping, Sequence, List
from contextlib import contextmanager

from . import (
    c2r_uninstall,
    constants,
    detect,
    elevation,
    logging_ext,
    msi_uninstall,
    fs_tools,
    registry_tools,
    tasks_services,
)


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


def uninstall_products(config: Mapping[str, object], *, dry_run: bool = False, retries: int | None = None) -> None:
    """!
    @brief Native Click-to-Run uninstall wrapper matching OffScrubC2R behavior.
    @details Reuses :mod:`c2r_uninstall` while preserving the OffScrub-style
    logging and option surface so callers can migrate directly.
    """

    human_logger = logging_ext.get_human_logger()
    human_logger.info("OffScrub native C2R: starting uninstall (dry_run=%s)", bool(dry_run))

    kwargs = {"dry_run": dry_run}
    if retries is not None:
        kwargs["retries"] = retries

    c2r_uninstall.uninstall_products(config, **kwargs)


def uninstall_msi_products(products: Sequence[Mapping[str, object] | str], *, dry_run: bool = False, retries: int | None = None) -> None:
    """!
    @brief Native MSI OffScrub entry point.
    @details Mirrors the semantics of the VBS MSI helpers by calling into
    :mod:`msi_uninstall` and preserving logging and retry semantics.
    """

    human_logger = logging_ext.get_human_logger()
    human_logger.info("OffScrub native MSI: uninstalling %d products (dry_run=%s)", len(list(products)), bool(dry_run))

    kwargs = {"dry_run": dry_run}
    if retries is not None:
        kwargs["retries"] = retries

    msi_uninstall.uninstall_products(products, **kwargs)


def _parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    """!
    @brief Argument parser that supports native and legacy OffScrub shapes.
    """

    parser = argparse.ArgumentParser(prog="off_scrub_native")
    sub = parser.add_subparsers(dest="command")

    c2r = sub.add_parser("c2r", help="Click-to-Run uninstall wrapper")
    c2r.add_argument("--dry-run", action="store_true", dest="dry_run")
    c2r.add_argument("--retries", type=int, default=None)
    c2r.add_argument("--release-ids", nargs="*", default=None)
    c2r.add_argument("--display-name", default=None)
    c2r.add_argument("legacy_args", nargs="*", help=argparse.SUPPRESS)

    msi = sub.add_parser("msi", help="MSI uninstall wrapper")
    msi.add_argument("--dry-run", action="store_true", dest="dry_run")
    msi.add_argument("--retries", type=int, default=None)
    msi.add_argument("--product-codes", nargs="*", default=None)
    msi.add_argument("legacy_args", nargs="*", help=argparse.SUPPRESS)

    return parser.parse_args(argv)


def _normalize_guid_token(token: str) -> str:
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


def _infer_version_group_from_script(script_path: Path | None, default: str | None = None) -> str | None:
    """!
    @brief Infer an OffScrub version group from the script filename.
    """

    if script_path is None:
        return default
    name = script_path.name.lower()
    return _SCRIPT_VERSION_HINTS.get(name, default)


def _parse_legacy_arguments(command: str, argv: Sequence[str]) -> LegacyInvocation:
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

        if upper == "ALL":
            flags["all"] = True
            index += 1
            continue

        stripped = upper[1:] if upper.startswith("/") else upper

        if stripped in ("PREVIEW", "DETECTONLY", "P"):
            flags["detect_only"] = True
            index += 1
            continue
        if stripped in ("QUIET", "Q", "QB", "QB-", "QND", "PASSIVE"):
            flags["quiet"] = True
            index += 1
            continue
        if stripped in ("NOREBOOT",):
            flags["no_reboot"] = True
            index += 1
            continue
        if stripped in ("NOCANCEL", "N"):
            flags["no_cancel"] = True
            index += 1
            continue
        if stripped in ("NOELEVATE", "NE"):
            flags["no_elevate"] = True
            index += 1
            continue
        if stripped in ("OFFLINE", "FORCEOFFLINE"):
            flags["offline"] = True
            index += 1
            continue
        if stripped in ("KEEPLICENSE", "KL"):
            flags["keep_license"] = True
            index += 1
            continue
        if stripped in ("KEEPUSERSETTINGS", "K"):
            flags["keep_user_settings"] = True
            index += 1
            continue
        if stripped in ("KEEPSOFTGRID", "KEEPSG"):
            flags["keep_softgrid"] = True
            index += 1
            continue
        if stripped in ("CLEARADDINREG", "CLEARADDINSREG"):
            flags["clear_addin_registry"] = True
            index += 1
            continue
        if stripped in ("DELETEUSERSETTINGS", "D"):
            flags["delete_user_settings"] = True
            index += 1
            continue
        if stripped in ("ENDCURRENTINSTALLS", "ECI"):
            flags["end_current_installs"] = True
            index += 1
            continue
        if stripped in ("REMOVELYNC",):
            flags["remove_lync"] = True
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
            index += 1
            continue

        if command == "msi" and _GUID_PATTERN.fullmatch(token):
            product_codes.append(_normalize_guid_token(token))
            index += 1
            continue
        if command == "c2r" and not token.startswith("/"):
            release_ids.append(token)
            index += 1
            continue

        unknown.append(token)
        index += 1

    version_group = _infer_version_group_from_script(
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


def _derive_execution_directives(legacy: LegacyInvocation, *, dry_run: bool) -> ExecutionDirectives:
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
    )


def _infer_msi_group(entry: Mapping[str, object]) -> str | None:
    """!
    @brief Infer MSI OffScrub version grouping from inventory metadata.
    """

    properties = entry.get("properties") if isinstance(entry, Mapping) else None
    supported = properties.get("supported_versions") if isinstance(properties, Mapping) else None
    if isinstance(supported, Sequence) and not isinstance(supported, (str, bytes)):
        for version in supported:
            group = constants.MSI_UNINSTALL_VERSION_GROUPS.get(str(version))
            if group:
                return group

    version_text = str(entry.get("version") or "").strip()
    if version_text:
        major = version_text.split(".", 1)[0]
        mapped = _MSI_MAJOR_VERSION_MAP.get(major)
        if mapped:
            return mapped
    return None


def _infer_c2r_group(entry: Mapping[str, object]) -> str | None:
    """!
    @brief Infer Click-to-Run grouping from inventory metadata.
    """

    properties = entry.get("properties") if isinstance(entry, Mapping) else None
    supported = properties.get("supported_versions") if isinstance(properties, Mapping) else None
    if isinstance(supported, Sequence) and not isinstance(supported, (str, bytes)):
        for version in supported:
            group = constants.C2R_UNINSTALL_VERSION_GROUPS.get(str(version))
            if group:
                return group
    return "c2r"


def _select_msi_targets(invocation: LegacyInvocation, inventory: Mapping[str, Any]) -> List[Mapping[str, object]]:
    """!
    @brief Filter MSI inventory entries to match the legacy invocation.
    """

    products: List[Mapping[str, object]] = []
    available = inventory.get("msi", []) if isinstance(inventory, Mapping) else []
    desired_codes = {code.upper() for code in invocation.product_codes if code}
    allow_all = bool(invocation.flags.get("all") or desired_codes)

    for entry in available:
        if not isinstance(entry, Mapping):
            continue
        product_code = _normalize_guid_token(str(entry.get("product_code") or ""))
        if not product_code:
            continue
        if desired_codes and product_code.upper() not in desired_codes:
            continue

        group = _infer_msi_group(entry)
        if invocation.version_group and group and group != invocation.version_group:
            continue
        if not allow_all and invocation.version_group is None:
            continue

        candidate = dict(entry)
        candidate["product_code"] = product_code
        products.append(candidate)

    if not products and desired_codes:
        products.extend({"product_code": code} for code in desired_codes)

    return products


def _select_c2r_targets(invocation: LegacyInvocation, inventory: Mapping[str, Any]) -> List[Mapping[str, object]]:
    """!
    @brief Filter Click-to-Run inventory entries to match the legacy invocation.
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
        if invocation.version_group and group and group != invocation.version_group:
            continue
        if not allow_all and invocation.version_group is None:
            continue

        targets.append(dict(entry))

    if not targets and desired_release_ids:
        targets.append({"release_ids": list(desired_release_ids)})

    return targets


def _perform_optional_cleanup(directives: ExecutionDirectives, *, dry_run: bool, kind: str | None = None) -> None:
    """!
    @brief Execute optional cleanup implied by legacy flags.
    @param kind Optional legacy command identifier (``c2r`` or ``msi``) used to scope cleanup.
    """

    human_logger = logging_ext.get_human_logger()

    if kind == "c2r":
        human_logger.info("Removing Click-to-Run scheduled tasks referenced by legacy scripts.")
        tasks_services.delete_tasks(constants.C2R_CLEANUP_TASKS, dry_run=dry_run)

    if not directives.skip_shortcut_detection:
        human_logger.info("Removing legacy Office shortcuts from known Start Menu roots.")
        fs_tools.remove_paths(_SHORTCUT_PATHS, dry_run=dry_run)
    else:
        human_logger.info("Skipping shortcut cleanup per legacy SkipSD flag.")

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


@contextmanager
def _quiet_logging(enabled: bool, human_logger: logging.Logger):
    """!
    @brief Temporarily raise the human logger threshold when quiet mode is set.
    """

    previous_level = human_logger.level
    if enabled:
        human_logger.setLevel(max(logging.WARNING, previous_level))
    try:
        yield
    finally:
        if enabled:
            human_logger.setLevel(previous_level)


def _log_flag_effects(
    legacy: LegacyInvocation, directives: ExecutionDirectives, human_logger
) -> None:
    """!
    @brief Emit informational logs describing how legacy flags are applied.
    """

    if directives.keep_license:
        human_logger.info(
            "Legacy keep-license flag set; license cleanup steps will be skipped if scheduled."
        )
    if directives.skip_shortcut_detection:
        human_logger.info(
            "Legacy SkipSD flag set; shortcut cleanup will be skipped in native mode."
        )
    if directives.offline:
        human_logger.info("Legacy offline flag set; Click-to-Run config will be marked offline.")
    if directives.reruns > 1:
        human_logger.info("Legacy test-rerun flag set; uninstall passes will run %d times.", directives.reruns)
    if directives.quiet:
        human_logger.info("Legacy quiet flag set; human log verbosity reduced for this run.")
    if directives.no_reboot:
        human_logger.info("Legacy no-reboot flag set; reboot recommendations will be suppressed.")
    if directives.delete_user_settings:
        human_logger.info("Legacy delete-user-settings flag set; user settings would be purged where applicable.")
    if directives.keep_user_settings:
        human_logger.info("Legacy keep-user-settings flag set; user settings cleanup will be skipped where applicable.")
    if directives.clear_addin_registry:
        human_logger.info("Legacy clear add-in registry flag set; add-in registry cleanup would be executed where applicable.")
    if directives.remove_vba:
        human_logger.info("Legacy remove VBA flag set; VBA-only package cleanup would be executed where applicable.")

    handled = {
        "all",
        "detect_only",
        "offline",
        "keep_license",
        "skip_shortcut_detection",
        "test_rerun",
        "quiet",
        "no_reboot",
        "delete_user_settings",
        "keep_user_settings",
        "clear_addin_registry",
        "remove_vba",
    }
    unmapped = sorted(
        flag for flag, enabled in legacy.flags.items() if enabled and flag not in handled
    )
    if unmapped:
        human_logger.info("Legacy flags not yet implemented in native flow: %s", ", ".join(unmapped))


def main(argv: Sequence[str] | None = None) -> int:
    """!
    @brief CLI entrypoint to mimic OffScrub script behaviour.
    @details Usage examples:
      - ``python -m office_janitor.off_scrub_native c2r --release-ids PRODUCTION``
      - ``python -m office_janitor.off_scrub_native msi --product-codes {GUID}``
    """

    argv_list = list(argv) if argv is not None else list(sys.argv[1:])
    args = _parse_args(argv_list)
    human_logger = logging_ext.get_human_logger()

    exit_code = 0
    try:
        if args.command == "c2r":
            legacy = _parse_legacy_arguments("c2r", getattr(args, "legacy_args", []) or [])
            if legacy.unknown:
                human_logger.warning("Unrecognised legacy arguments: %s", ", ".join(legacy.unknown))
            if legacy.log_directory:
                human_logger.info("Legacy log directory requested: %s", legacy.log_directory)

            if legacy.flags.get("no_elevate") and elevation.is_admin() and not os.environ.get("OFFICE_JANITOR_DEELEVATED"):
                human_logger.info("Legacy no-elevate flag set; re-launching OffScrub native as limited user.")
                result = elevation.run_as_limited_user(
                    [sys.executable, "-m", "office_janitor.off_scrub_native", *argv_list],
                    event="off_scrub_deelevate",
                    env_overrides={"OFFICE_JANITOR_DEELEVATED": "1"},
                )
                return result.returncode

            inventory = detect.gather_office_inventory()
            if args.release_ids:
                legacy.release_ids.extend([rid for rid in args.release_ids if rid])

            targets = _select_c2r_targets(legacy, inventory)
            if not targets and not legacy.release_ids:
                human_logger.info("No Click-to-Run installations detected for legacy request.")
                return 0

            dry_run = bool(args.dry_run or legacy.flags.get("detect_only"))
            directives = _derive_execution_directives(legacy, dry_run=dry_run)

            with _quiet_logging(directives.quiet, human_logger), tasks_services.suppress_reboot_recommendations(directives.no_reboot):
                _log_flag_effects(legacy, directives, human_logger)
                for target in targets:
                    if args.display_name and isinstance(target, MutableMapping):
                        merged = dict(target)
                        merged["product"] = args.display_name
                        target = merged
                    if directives.offline and isinstance(target, MutableMapping):
                        target = dict(target)
                        target["offline"] = True
                    if directives.keep_license and isinstance(target, MutableMapping):
                        target = dict(target)
                        target["keep_license"] = True
                    for attempt in range(1, directives.reruns + 1):
                        if directives.reruns > 1:
                            human_logger.info("Legacy rerun pass %d/%d for Click-to-Run target.", attempt, directives.reruns)
                        uninstall_products(target, dry_run=dry_run, retries=args.retries)
            _perform_optional_cleanup(directives, dry_run=dry_run, kind="c2r")
            exit_code = 0
        elif args.command == "msi":
            legacy = _parse_legacy_arguments("msi", getattr(args, "legacy_args", []) or [])
            if legacy.unknown:
                human_logger.warning("Unrecognised legacy arguments: %s", ", ".join(legacy.unknown))
            if legacy.log_directory:
                human_logger.info("Legacy log directory requested: %s", legacy.log_directory)

            if legacy.flags.get("no_elevate") and elevation.is_admin() and not os.environ.get("OFFICE_JANITOR_DEELEVATED"):
                human_logger.info("Legacy no-elevate flag set; re-launching OffScrub native as limited user.")
                result = elevation.run_as_limited_user(
                    [sys.executable, "-m", "office_janitor.off_scrub_native", *argv_list],
                    event="off_scrub_deelevate",
                    env_overrides={"OFFICE_JANITOR_DEELEVATED": "1"},
                )
                return result.returncode

            inventory = detect.gather_office_inventory()
            if args.product_codes:
                legacy.product_codes.extend([code for code in args.product_codes if code])

            selected_products = _select_msi_targets(legacy, inventory)
            if not selected_products:
                human_logger.info("No MSI installations matched the legacy OffScrub request.")
                return 0

            dry_run = bool(args.dry_run or legacy.flags.get("detect_only"))
            directives = _derive_execution_directives(legacy, dry_run=dry_run)
            with _quiet_logging(directives.quiet, human_logger), tasks_services.suppress_reboot_recommendations(directives.no_reboot):
                _log_flag_effects(legacy, directives, human_logger)
                products_to_use: List[Mapping[str, object]] = []
                for entry in selected_products:
                    if not isinstance(entry, MutableMapping):
                        products_to_use.append({"product_code": entry})
                        continue
                    merged: MutableMapping[str, object] = dict(entry)
                    if directives.delete_user_settings:
                        merged["delete_user_settings"] = True
                    if directives.keep_user_settings:
                        merged["keep_user_settings"] = True
                    if directives.clear_addin_registry:
                        merged["clear_addin_registry"] = True
                    if directives.remove_vba:
                        merged["remove_vba"] = True
                    products_to_use.append(merged)
                for attempt in range(1, directives.reruns + 1):
                    if directives.reruns > 1:
                        human_logger.info("Legacy rerun pass %d/%d for MSI targets.", attempt, directives.reruns)
                    uninstall_msi_products(products_to_use, dry_run=dry_run, retries=args.retries)
            _perform_optional_cleanup(directives, dry_run=dry_run, kind="msi")
            exit_code = 0
        else:
            human_logger.info("No command supplied; nothing to do.")
            exit_code = 2
    except Exception as exc:  # pragma: no cover - propagate to caller
        human_logger.error("OffScrub native operation failed: %s", exc)
        exit_code = 1

    # Map pending reboot recommendations into a legacy-style return code bitmask.
    pending_reboots = tasks_services.consume_reboot_recommendations()
    if pending_reboots:
        human_logger.info("Reboot recommended for services: %s", ", ".join(pending_reboots))
        if exit_code == 0:
            exit_code = 2
        else:
            exit_code = exit_code | 2

    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
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
)
