"""!
@file main_plan.py
@brief Plan options collection utilities for Office Janitor.
@details Handles configuration file loading and CLI argument translation
to planning options with proper precedence.
"""

from __future__ import annotations

import json
import pathlib
import sys
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    import argparse

__all__ = [
    "load_config_file",
    "collect_plan_options",
]


def load_config_file(config_path: str | None) -> dict[str, object]:
    """!
    @brief Load and parse a JSON configuration file.
    @param config_path Path to the JSON config file, or None to skip.
    @returns Dictionary of configuration options, empty if no file specified.
    @raises SystemExit if the file cannot be read or parsed.
    """
    if not config_path:
        return {}

    path = pathlib.Path(config_path).expanduser().resolve()
    if not path.exists():
        print(f"Error: Configuration file not found: {path}", file=sys.stderr)
        raise SystemExit(1)

    try:
        with open(path, encoding="utf-8") as f:
            config = json.load(f)
        if not isinstance(config, dict):
            print(
                f"Error: Configuration file must contain a JSON object: {path}",
                file=sys.stderr,
            )
            raise SystemExit(1)
        return config
    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON in configuration file: {path}\n{e}", file=sys.stderr)
        raise SystemExit(1) from e
    except OSError as e:
        print(f"Error: Cannot read configuration file: {path}\n{e}", file=sys.stderr)
        raise SystemExit(1) from e


def collect_plan_options(args: argparse.Namespace, mode: str) -> dict[str, object]:
    """!
    @brief Translate parsed CLI arguments into planning options.
    @details Options are resolved with the following precedence (highest first):
    1. CLI arguments explicitly specified
    2. JSON config file values (if --config provided)
    3. Built-in defaults
    @param args Parsed command-line arguments.
    @param mode Mode string for the operation.
    @returns Dictionary of planning options.
    """
    # Load config file if specified
    config = load_config_file(getattr(args, "config", None))

    # Helper to get value with precedence: CLI > config > default
    def _get(
        attr: str,
        default: object = None,
        config_key: str | None = None,
        is_bool: bool = False,
    ) -> object:
        """Get option value with CLI > config > default precedence."""
        cli_val = getattr(args, attr, None)
        cfg_key = config_key or attr.replace("_", "-")  # CLI uses underscores, JSON uses hyphens

        # For boolean flags, CLI sets True explicitly; check if user specified it
        if is_bool:
            if cli_val:
                return True
            if cfg_key in config:
                return bool(config[cfg_key])
            return bool(default) if default else False

        # For non-boolean, None means "not specified"
        if cli_val is not None:
            return cli_val
        if cfg_key in config:
            return config[cfg_key]
        return default

    # Resolve max_passes:
    # --registry-only or --skip-uninstall (=0) > --passes > --max-passes > config > default (1)
    skip_uninstall = _get("skip_uninstall", False, "skip-uninstall", is_bool=True)
    registry_only = _get("registry_only", False, "registry-only", is_bool=True)
    cli_passes = getattr(args, "passes", None)
    cli_max_passes = getattr(args, "max_passes", None)
    config_passes = config.get("passes") or config.get("max-passes")
    resolved_passes = (
        0
        if (skip_uninstall or registry_only)
        else (cli_passes or cli_max_passes or config_passes or 1)
    )

    options: dict[str, object] = {
        # Mode & core
        "mode": mode,
        "dry_run": _get("dry_run", False, is_bool=True),
        "force": _get("force", False, is_bool=True),
        "yes": _get("yes", False, is_bool=True),
        "include": _get("include", None),
        "target": _get("target", None),
        "diagnose": _get("diagnose", False, is_bool=True),
        "cleanup_only": _get("cleanup_only", False, "cleanup-only", is_bool=True),
        "auto_all": _get("auto_all", False, "auto-all", is_bool=True),
        "allow_unsupported_windows": _get("allow_unsupported_windows", False, is_bool=True),
        # Uninstall method
        "uninstall_method": _get("uninstall_method", "auto"),
        "force_app_shutdown": _get("force_app_shutdown", False, is_bool=True),
        "no_force_app_shutdown": _get("no_force_app_shutdown", False, is_bool=True),
        "product_codes": _get("product_codes", None),
        "release_ids": _get("release_ids", None),
        # Scrubbing
        "scrub_level": _get("scrub_level", "standard"),
        "max_passes": int(resolved_passes),
        "skip_processes": _get("skip_processes", False, is_bool=True) or registry_only,
        "skip_services": _get("skip_services", False, is_bool=True) or registry_only,
        "skip_tasks": _get("skip_tasks", False, is_bool=True),
        "skip_registry": _get("skip_registry", False, is_bool=True),
        "skip_filesystem": _get("skip_filesystem", False, is_bool=True) or registry_only,
        "registry_only": registry_only,
        "clean_msocache": _get("clean_msocache", False, is_bool=True),
        "clean_appx": _get("clean_appx", False, is_bool=True),
        "clean_wi_metadata": _get("clean_wi_metadata", False, is_bool=True),
        # License & activation
        # Restore point: enabled by default unless --no-restore-point is specified
        # or explicitly enabled with --restore-point/--create-restore-point
        "create_restore_point": (
            _get("create_restore_point", False, is_bool=True)
            or not _get("no_restore_point", False, is_bool=True)
        ),
        "no_license": (
            _get("no_license", False, is_bool=True)
            or _get("keep_license", False, is_bool=True)
            or registry_only
        ),
        "keep_license": _get("keep_license", False, is_bool=True),
        "clean_spp": _get("clean_spp", False, is_bool=True),
        "clean_ospp": _get("clean_ospp", False, is_bool=True),
        "clean_vnext": _get("clean_vnext", False, is_bool=True),
        "clean_all_licenses": _get("clean_all_licenses", False, is_bool=True),
        # User data
        "keep_templates": _get("keep_templates", False, is_bool=True),
        "keep_user_settings": _get("keep_user_settings", False, is_bool=True),
        "delete_user_settings": _get("delete_user_settings", False, is_bool=True),
        "keep_outlook_data": _get("keep_outlook_data", False, is_bool=True),
        "keep_outlook_signatures": _get("keep_outlook_signatures", False, is_bool=True),
        "clean_shortcuts": _get("clean_shortcuts", False, is_bool=True),
        "skip_shortcut_detection": _get("skip_shortcut_detection", False, is_bool=True),
        # Registry cleanup
        "clean_addin_registry": _get("clean_addin_registry", False, is_bool=True),
        "clean_com_registry": _get("clean_com_registry", False, is_bool=True),
        "clean_shell_extensions": _get("clean_shell_extensions", False, is_bool=True),
        "clean_typelibs": _get("clean_typelibs", False, is_bool=True),
        "clean_protocol_handlers": _get("clean_protocol_handlers", False, is_bool=True),
        "remove_vba": _get("remove_vba", False, is_bool=True),
        # Output & paths
        "timeout": _get("timeout", None),
        "backup": _get("backup", None),
        "verbose": _get("verbose", 0),
        # Retry & resilience
        "retries": _get("retries", 4),
        "retry_delay": _get("retry_delay", 3),
        "retry_delay_max": _get("retry_delay_max", 30),
        "no_reboot": _get("no_reboot", False, is_bool=True),
        "offline": _get("offline", False, is_bool=True),
        # Advanced
        "skip_preflight": _get("skip_preflight", False, is_bool=True),
        "skip_backup": _get("skip_backup", False, is_bool=True),
        "skip_verification": _get("skip_verification", False, is_bool=True),
        "schedule_reboot": _get("schedule_reboot", False, is_bool=True),
        "no_schedule_delete": _get("no_schedule_delete", False, is_bool=True),
        "msiexec_args": _get("msiexec_args", None),
        "c2r_args": _get("c2r_args", None),
        "odt_args": _get("odt_args", None),
        # OffScrub legacy
        "offscrub_all": _get("offscrub_all", False, is_bool=True),
        "offscrub_ose": _get("offscrub_ose", False, is_bool=True),
        "offscrub_offline": _get("offscrub_offline", False, is_bool=True),
        "offscrub_quiet": _get("offscrub_quiet", False, is_bool=True),
        "offscrub_test_rerun": _get("offscrub_test_rerun", False, is_bool=True),
        "offscrub_bypass": _get("offscrub_bypass", False, is_bool=True),
        "offscrub_fast_remove": _get("offscrub_fast_remove", False, is_bool=True),
        "offscrub_scan_components": _get("offscrub_scan_components", False, is_bool=True),
        "offscrub_return_error": _get("offscrub_return_error", False, is_bool=True),
        # Repair options
        "repair_timeout": _get("repair_timeout", 3600),
        # Miscellaneous
        "limited_user": _get("limited_user", False, is_bool=True),
    }

    # Auto-all mode enables FULL scrubbing like the VBS scripts:
    # - All Office versions targeted
    # - All cleanup options enabled
    # - Force app shutdown enabled
    # - All license cleaning enabled
    # Only apply auto-all defaults for options the user didn't explicitly specify
    if mode == "auto-all":
        # Auto-all defaults (only applied if user didn't explicitly specify different)
        auto_all_defaults: dict[str, object] = {
            # Aggressive scrub level (if user didn't specify a different level)
            "scrub_level": "nuclear",
            # Force close apps
            "force_app_shutdown": True,
            "force": True,
            # Clean everything
            "clean_msocache": True,
            "clean_appx": True,
            "clean_wi_metadata": True,
            # Clean all licenses
            "clean_spp": True,
            "clean_ospp": True,
            "clean_vnext": True,
            "clean_all_licenses": True,
            # Clean registry
            "clean_addin_registry": True,
            "clean_com_registry": True,
            "clean_shell_extensions": True,
            "clean_typelibs": True,
            "clean_protocol_handlers": True,
            # Clean user data
            "clean_shortcuts": True,
            "delete_user_settings": True,
            # VBA
            "remove_vba": True,
            # OffScrub-style options
            "offscrub_all": True,
            "offscrub_ose": True,
        }
        # For scrub_level, only override if user used default "standard"
        if options.get("scrub_level") == "standard":
            options["scrub_level"] = auto_all_defaults["scrub_level"]
        # For boolean flags, enable if not already True (boolean options default False)
        for key, value in auto_all_defaults.items():
            if key == "scrub_level":
                continue  # Already handled above
            # Only set True if user didn't explicitly enable (these default to False)
            if not options.get(key):
                options[key] = value

    return options
