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
from collections.abc import Mapping, MutableMapping, Sequence
from contextlib import contextmanager

from . import (
    c2r_uninstall,
    detect,
    elevation,
    logging_ext,
    msi_uninstall,
    tasks_services,
)
from . import off_scrub_helpers as _helpers
from .off_scrub_helpers import ExecutionDirectives, LegacyInvocation

fs_tools = _helpers.fs_tools
registry_tools = _helpers.registry_tools
constants = _helpers.constants


def _wrap_parse_legacy_arguments(command: str, argv: Sequence[str]) -> LegacyInvocation:
    """!
    @brief Compatibility wrapper for tests and legacy entrypoints.
    """

    return _helpers.parse_legacy_arguments(command, argv)


def _wrap_select_c2r_targets(
    invocation: LegacyInvocation, inventory: Mapping[str, object]
) -> list[Mapping[str, object]]:
    return _helpers.select_c2r_targets(invocation, inventory)


def _wrap_select_msi_targets(
    invocation: LegacyInvocation, inventory: Mapping[str, object]
) -> list[Mapping[str, object]]:
    return _helpers.select_msi_targets(invocation, inventory)


def _wrap_perform_optional_cleanup(
    directives: ExecutionDirectives, *, dry_run: bool, kind: str | None = None
) -> None:
    # Keep helper module references in sync with monkeypatches applied to this module during tests.
    _helpers.fs_tools = fs_tools
    _helpers.registry_tools = registry_tools
    _helpers.tasks_services = tasks_services
    return _helpers.perform_optional_cleanup(directives, dry_run=dry_run, kind=kind)


# Public-facing aliases kept for tests/backwards compatibility
_parse_legacy_arguments = _wrap_parse_legacy_arguments
_derive_execution_directives = _helpers.derive_execution_directives
_select_c2r_targets = _wrap_select_c2r_targets
_select_msi_targets = _wrap_select_msi_targets
_perform_optional_cleanup = _wrap_perform_optional_cleanup


def uninstall_products(
    config: Mapping[str, object], *, dry_run: bool = False, retries: int | None = None
) -> None:
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


def uninstall_msi_products(
    products: Sequence[Mapping[str, object] | str],
    *,
    dry_run: bool = False,
    retries: int | None = None,
) -> None:
    """!
    @brief Native MSI OffScrub entry point.
    @details Mirrors the semantics of the VBS MSI helpers by calling into
    :mod:`msi_uninstall` and preserving logging and retry semantics.
    """

    human_logger = logging_ext.get_human_logger()
    human_logger.info(
        "OffScrub native MSI: uninstalling %d products (dry_run=%s)",
        len(list(products)),
        bool(dry_run),
    )

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
        human_logger.info(
            "Legacy test-rerun flag set; uninstall passes will run %d times.", directives.reruns
        )
    if directives.quiet:
        human_logger.info("Legacy quiet flag set; human log verbosity reduced for this run.")
    if directives.no_reboot:
        human_logger.info("Legacy no-reboot flag set; reboot recommendations will be suppressed.")
    if directives.delete_user_settings:
        human_logger.info(
            "Legacy delete-user-settings flag set; user settings would be purged where applicable."
        )
    if directives.keep_user_settings:
        human_logger.info(
            "Legacy keep-user-settings flag set; user settings cleanup will be skipped where "
            "applicable."
        )
    if directives.clear_addin_registry:
        human_logger.info(
            "Legacy clear add-in registry flag set; add-in registry cleanup would be executed "
            "where applicable."
        )
    if directives.remove_vba:
        human_logger.info(
            "Legacy remove VBA flag set; VBA-only package cleanup would be executed where "
            "applicable."
        )
    if directives.return_error_or_success:
        human_logger.info(
            "Legacy ReturnErrorOrSuccess flag set; exit code will be reduced to success unless a "
            "reboot bit is present."
        )

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
        "return_error_or_success",
    }
    unmapped = sorted(
        flag for flag, enabled in legacy.flags.items() if enabled and flag not in handled
    )
    if unmapped:
        human_logger.info(
            "Legacy flags not yet implemented in native flow: %s", ", ".join(unmapped)
        )


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
    machine_logger = logging_ext.get_machine_logger()

    exit_code = 0
    return_success_on_error = False
    try:
        if args.command == "c2r":
            legacy = _parse_legacy_arguments("c2r", getattr(args, "legacy_args", []) or [])
            if legacy.unknown:
                human_logger.warning("Unrecognised legacy arguments: %s", ", ".join(legacy.unknown))
            if legacy.log_directory:
                human_logger.info("Legacy log directory requested: %s", legacy.log_directory)

            if (
                legacy.flags.get("no_elevate")
                and elevation.is_admin()
                and not os.environ.get("OFFICE_JANITOR_DEELEVATED")
            ):
                human_logger.info(
                    "Legacy no-elevate flag set; re-launching OffScrub native as limited user."
                )
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
            return_success_on_error = directives.return_error_or_success

            with (
                _quiet_logging(directives.quiet, human_logger),
                tasks_services.suppress_reboot_recommendations(directives.no_reboot),
            ):
                _log_flag_effects(legacy, directives, human_logger)
                machine_logger.info(
                    "stage0_detection", extra={"event": "stage0_detection", "command": "c2r"}
                )
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
                            human_logger.info(
                                "Legacy rerun pass %d/%d for Click-to-Run target.",
                                attempt,
                                directives.reruns,
                            )
                        machine_logger.info(
                            "stage1_uninstall",
                            extra={
                                "event": "stage1_uninstall",
                                "command": "c2r",
                                "attempt": attempt,
                                "attempts": directives.reruns,
                            },
                        )
                        uninstall_products(target, dry_run=dry_run, retries=args.retries)
            with _quiet_logging(directives.quiet, human_logger):
                _perform_optional_cleanup(directives, dry_run=dry_run, kind="c2r")
            exit_code = 0
        elif args.command == "msi":
            legacy = _parse_legacy_arguments("msi", getattr(args, "legacy_args", []) or [])
            if legacy.unknown:
                human_logger.warning("Unrecognised legacy arguments: %s", ", ".join(legacy.unknown))
            if legacy.log_directory:
                human_logger.info("Legacy log directory requested: %s", legacy.log_directory)

            if (
                legacy.flags.get("no_elevate")
                and elevation.is_admin()
                and not os.environ.get("OFFICE_JANITOR_DEELEVATED")
            ):
                human_logger.info(
                    "Legacy no-elevate flag set; re-launching OffScrub native as limited user."
                )
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
            return_success_on_error = directives.return_error_or_success
            with (
                _quiet_logging(directives.quiet, human_logger),
                tasks_services.suppress_reboot_recommendations(directives.no_reboot),
            ):
                _log_flag_effects(legacy, directives, human_logger)
                machine_logger.info(
                    "stage0_detection", extra={"event": "stage0_detection", "command": "msi"}
                )
                products_to_use: list[Mapping[str, object]] = []
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
                        human_logger.info(
                            "Legacy rerun pass %d/%d for MSI targets.", attempt, directives.reruns
                        )
                    machine_logger.info(
                        "stage1_uninstall",
                        extra={
                            "event": "stage1_uninstall",
                            "command": "msi",
                            "attempt": attempt,
                            "attempts": directives.reruns,
                        },
                    )
                    uninstall_msi_products(products_to_use, dry_run=dry_run, retries=args.retries)
            with _quiet_logging(directives.quiet, human_logger):
                _perform_optional_cleanup(directives, dry_run=dry_run, kind="msi")
            exit_code = 0
        else:
            human_logger.info("No command supplied; nothing to do.")
            exit_code = 2
    except Exception as exc:  # pragma: no cover - propagate to caller
        human_logger.error("OffScrub native operation failed: %s", exc)
        exit_code = 1

    if return_success_on_error and exit_code not in (0, 2):
        exit_code = 2 if exit_code & 2 else 0

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
