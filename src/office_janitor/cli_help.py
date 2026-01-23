"""!
@brief CLI help text and argument definitions for Office Janitor.
@details This module centralizes all CLI argument parsing and help text generation,
enabling a streamlined help system and consistent command-line interface.
"""

from __future__ import annotations

import argparse
import os
import sys
from typing import TYPE_CHECKING, Any

# Import subcommand-specific options
from .cli_c2r import C2R_EPILOG, add_c2r_subcommand_options
from .cli_config import CONFIG_EPILOG, add_config_subcommand_options
from .cli_diagnose import DIAGNOSE_EPILOG, add_diagnose_subcommand_options
from .cli_install import INSTALL_EPILOG, add_install_subcommand_options
from .cli_license import LICENSE_EPILOG, add_license_subcommand_options
from .cli_odt import ODT_EPILOG, add_odt_subcommand_options
from .cli_offscrub import OFFSCRUB_EPILOG, add_offscrub_subcommand_options
from .cli_remove import REMOVE_EPILOG, add_remove_subcommand_options
from .cli_repair import REPAIR_EPILOG, add_repair_subcommand_options

if TYPE_CHECKING:
    pass


def _should_pause_on_exit() -> bool:
    """!
    @brief Determine if we should pause before exit.
    @details Only pause when running interactively (stdin is a TTY) and not in tests.
    @returns True if pause is appropriate.
    """
    # Don't pause in test environments
    if os.environ.get("PYTEST_CURRENT_TEST"):
        return False
    # Don't pause if stdin is not a TTY (e.g., piped input)
    if not sys.stdin.isatty():
        return False
    return True


# ---------------------------------------------------------------------------
# Custom Argparse Actions with Pause
# ---------------------------------------------------------------------------


class HelpActionWithPause(argparse.Action):
    """!
    @brief Custom help action that pauses before exiting.
    @details Prevents the console window from closing immediately after
    displaying help when run from a GUI shortcut or double-click.
    Only pauses when stdin is a TTY (not in tests or piped input).
    """

    def __init__(
        self,
        option_strings: list[str],
        dest: str = argparse.SUPPRESS,
        default: str = argparse.SUPPRESS,
        help: str | None = "show this help message and exit",  # noqa: A002
    ) -> None:
        super().__init__(
            option_strings=option_strings,
            dest=dest,
            default=default,
            nargs=0,
            help=help,
        )

    def __call__(
        self,
        parser: argparse.ArgumentParser,
        namespace: argparse.Namespace,
        values: Any,
        option_string: str | None = None,
    ) -> None:
        parser.print_help()
        if _should_pause_on_exit():
            print("\nPress Enter to exit...")
            try:
                input()
            except (EOFError, OSError, KeyboardInterrupt):
                pass
        parser.exit()


class VersionActionWithPause(argparse.Action):
    """!
    @brief Custom version action that pauses before exiting.
    @details Prevents the console window from closing immediately after
    displaying version when run from a GUI shortcut or double-click.
    Only pauses when stdin is a TTY (not in tests or piped input).
    """

    def __init__(
        self,
        option_strings: list[str],
        version: str | None = None,
        dest: str = argparse.SUPPRESS,
        default: str = argparse.SUPPRESS,
        help: str | None = "show program's version number and exit",  # noqa: A002
    ) -> None:
        super().__init__(
            option_strings=option_strings,
            dest=dest,
            default=default,
            nargs=0,
            help=help,
        )
        self.version = version

    def __call__(
        self,
        parser: argparse.ArgumentParser,
        namespace: argparse.Namespace,
        values: Any,
        option_string: str | None = None,
    ) -> None:
        version = self.version
        if version:
            formatter = parser._get_formatter()
            formatter.add_text(version)
            parser._print_message(formatter.format_help(), sys.stdout)
        if _should_pause_on_exit():
            print("\nPress Enter to exit...")
            try:
                input()
            except (EOFError, OSError, KeyboardInterrupt):
                pass
        parser.exit()


# ---------------------------------------------------------------------------
# Version & Program Info
# ---------------------------------------------------------------------------

PROGRAM_NAME = "office-janitor"
PROGRAM_DESCRIPTION = """\
Microsoft Office installation manager.

Commands:
  install    Deploy Office via ODT presets or custom configurations
  repair     Fix broken Office installations (quick or full repair)
  remove     Uninstall Office and clean up residual artifacts
  diagnose   Detect and report Office installations
  odt        Build and manage ODT XML configurations
  offscrub   OffScrub-style deep removal operations
  c2r        Direct Click-to-Run passthrough operations
  license    Manage Office licensing and activation
  config     Generate and manage configuration files

Run 'office-janitor <command> --help' for command-specific options.
"""

# ---------------------------------------------------------------------------
# Help Text Constants
# ---------------------------------------------------------------------------

EPILOG_TEXT = """\
EXAMPLES:
  office-janitor install --preset 365-proplus-x64
  office-janitor repair --quick
  office-janitor remove --dry-run
  office-janitor diagnose --plan report.json

LEGACY FLAGS:
  --auto-all, --auto-repair, --repair quick (still supported)

Docs: https://github.com/supermarsx/office-janitor
"""

# ---------------------------------------------------------------------------
# Argument Group Definitions
# ---------------------------------------------------------------------------


def add_mode_arguments(parser: argparse.ArgumentParser) -> argparse.ArgumentParser:
    """!
    @brief Add mode selection arguments (mutually exclusive).
    @param parser The ArgumentParser to add arguments to.
    @returns The parser for chaining.
    """
    modes = parser.add_mutually_exclusive_group()
    modes.add_argument(
        "--auto-all",
        action="store_true",
        help="Run full detection and scrub of all Office installations.",
    )
    modes.add_argument(
        "--target",
        metavar="VER",
        help="Target a specific Office version (2003-2024/365).",
    )
    modes.add_argument(
        "--diagnose",
        action="store_true",
        help="Emit inventory and plan without making changes.",
    )
    modes.add_argument(
        "--cleanup-only",
        action="store_true",
        help="Skip uninstalls; clean residue and licensing only.",
    )
    modes.add_argument(
        "--repair",
        choices=["quick", "full"],
        metavar="TYPE",
        help="Repair Office C2R (quick=local, full=online CDN).",
    )
    modes.add_argument(
        "--repair-config",
        metavar="XML",
        help="Repair/reconfigure using a custom XML configuration file.",
    )
    modes.add_argument(
        "--auto-repair",
        action="store_true",
        help="Auto-detect and repair all Office installations.",
    )
    return parser


def add_global_options(parser: argparse.ArgumentParser) -> argparse.ArgumentParser:
    """!
    @brief Add truly global options that apply to all commands.
    @param parser The ArgumentParser to add arguments to.
    @returns The parser for chaining.
    @details These are options that apply regardless of which command is used.
    Subcommand-specific options should be shown only in their respective help.
    """
    global_opts = parser.add_argument_group("Global Options")
    global_opts.add_argument(
        "--dry-run",
        "-n",
        action="store_true",
        help="Simulate actions without modifying the system.",
    )
    global_opts.add_argument(
        "--yes",
        "-y",
        action="store_true",
        help="Skip confirmation prompts (assume yes).",
    )
    global_opts.add_argument(
        "--config",
        "-c",
        metavar="JSON",
        help="Load options from a JSON configuration file.",
    )
    global_opts.add_argument(
        "--verbose",
        "-v",
        action="count",
        default=0,
        help="Increase output verbosity (-v, -vv, -vvv).",
    )
    global_opts.add_argument(
        "--quiet",
        "-q",
        action="store_true",
        help="Minimal console output (errors only).",
    )
    global_opts.add_argument(
        "--json",
        action="store_true",
        help="Mirror structured events to stdout.",
    )
    global_opts.add_argument(
        "--tui",
        action="store_true",
        help="Force the interactive text UI mode.",
    )
    global_opts.add_argument(
        "--no-color",
        action="store_true",
        help="Disable ANSI color codes.",
    )
    return parser


def add_legacy_options(parser: argparse.ArgumentParser) -> argparse.ArgumentParser:
    """!
    @brief Add legacy options for backward compatibility (hidden from help).
    @param parser The ArgumentParser to add arguments to.
    @returns The parser for chaining.
    @details These options are added for backward compatibility with older scripts
    and CLI usage patterns. They are suppressed from help output to keep the
    main help clean. Users should use `<command> --help` for detailed options.
    """
    # We add all the option groups but with help=SUPPRESS to hide from main help.
    # This maintains backward compatibility while keeping the main help clean.

    # Core Options (subset not already in global)
    parser.add_argument("--include", metavar="COMPONENTS", help=argparse.SUPPRESS)
    parser.add_argument("--force", "-f", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument(
        "--allow-unsupported-windows", action="store_true", help=argparse.SUPPRESS
    )
    parser.add_argument("--passes", type=int, default=None, metavar="N", help=argparse.SUPPRESS)

    # Repair Options
    parser.add_argument("--repair-odt", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--repair-c2r", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--repair-culture", metavar="LANG", default="en-us", help=argparse.SUPPRESS)
    parser.add_argument(
        "--repair-platform", choices=["x86", "x64"], metavar="ARCH", help=argparse.SUPPRESS
    )
    parser.add_argument("--repair-visible", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument(
        "--repair-timeout", type=int, default=3600, metavar="SEC", help=argparse.SUPPRESS
    )
    parser.add_argument("--repair-all-products", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--repair-preset", metavar="NAME", help=argparse.SUPPRESS)

    # Uninstall Method Options
    parser.add_argument(
        "--uninstall-method",
        choices=["auto", "msi", "c2r", "odt", "offscrub"],
        default=None,
        metavar="METHOD",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--msi-only",
        action="store_const",
        const="msi",
        dest="uninstall_method",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--c2r-only",
        action="store_const",
        const="c2r",
        dest="uninstall_method",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--use-odt",
        action="store_const",
        const="odt",
        dest="uninstall_method",
        help=argparse.SUPPRESS,
    )
    parser.add_argument("--force-app-shutdown", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--no-force-app-shutdown", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument(
        "--product-code",
        metavar="GUID",
        action="append",
        dest="product_codes",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--release-id",
        metavar="ID",
        action="append",
        dest="release_ids",
        help=argparse.SUPPRESS,
    )

    # Scrubbing Options
    parser.add_argument(
        "--scrub-level",
        choices=["minimal", "standard", "aggressive", "nuclear"],
        default=None,
        metavar="LEVEL",
        help=argparse.SUPPRESS,
    )
    parser.add_argument("--max-passes", type=int, default=None, metavar="N", help=argparse.SUPPRESS)
    parser.add_argument("--skip-uninstall", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--skip-processes", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--skip-services", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--skip-tasks", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--skip-registry", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--skip-filesystem", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--registry-only", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--clean-msocache", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--clean-appx", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--clean-wi-metadata", action="store_true", help=argparse.SUPPRESS)

    # License & Activation Options
    restore_point = parser.add_mutually_exclusive_group()
    restore_point.add_argument(
        "--restore-point",
        "--create-restore-point",
        action="store_true",
        dest="create_restore_point",
        help=argparse.SUPPRESS,
    )
    restore_point.add_argument(
        "--no-restore-point", action="store_true", help=argparse.SUPPRESS
    )
    parser.add_argument("--no-license", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--keep-license", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--clean-spp", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--clean-ospp", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--clean-vnext", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--clean-all-licenses", action="store_true", help=argparse.SUPPRESS)

    # User Data Options
    parser.add_argument("--keep-templates", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--keep-user-settings", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--delete-user-settings", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--keep-outlook-data", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--keep-outlook-signatures", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--clean-shortcuts", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--skip-shortcut-detection", action="store_true", help=argparse.SUPPRESS)

    # Registry Cleanup Options
    parser.add_argument("--clean-addin-registry", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--clean-com-registry", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--clean-shell-extensions", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--clean-typelibs", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--clean-protocol-handlers", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--remove-vba", action="store_true", help=argparse.SUPPRESS)

    # Output & Logging Options (subset not already in global)
    parser.add_argument("--plan", metavar="OUT", help=argparse.SUPPRESS)
    parser.add_argument("--logdir", metavar="DIR", help=argparse.SUPPRESS)
    parser.add_argument("--backup", metavar="DIR", help=argparse.SUPPRESS)
    parser.add_argument("--timeout", metavar="SEC", type=int, help=argparse.SUPPRESS)

    # TUI Options (subset not already in global)
    parser.add_argument("--tui-compact", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--tui-refresh", metavar="MS", type=int, help=argparse.SUPPRESS)
    parser.add_argument("--limited-user", action="store_true", help=argparse.SUPPRESS)

    # Retry & Resilience Options
    parser.add_argument("--retries", type=int, default=None, metavar="N", help=argparse.SUPPRESS)
    parser.add_argument("--retry-delay", type=int, default=3, metavar="SEC", help=argparse.SUPPRESS)
    parser.add_argument(
        "--retry-delay-max", type=int, default=30, metavar="SEC", help=argparse.SUPPRESS
    )
    parser.add_argument("--no-reboot", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--offline", action="store_true", help=argparse.SUPPRESS)

    # ODT Build Options
    parser.add_argument("--odt-build", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--odt-install", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--odt-preset", metavar="NAME", help=argparse.SUPPRESS)
    parser.add_argument(
        "--odt-product", metavar="ID", action="append", dest="odt_products", help=argparse.SUPPRESS
    )
    parser.add_argument(
        "--odt-language",
        metavar="CODE",
        action="append",
        dest="odt_languages",
        help=argparse.SUPPRESS,
    )
    parser.add_argument(
        "--odt-arch", choices=["32", "64"], default="64", metavar="BITS", help=argparse.SUPPRESS
    )
    parser.add_argument("--odt-channel", metavar="CHANNEL", help=argparse.SUPPRESS)
    parser.add_argument("--odt-output", metavar="FILE", help=argparse.SUPPRESS)
    parser.add_argument("--odt-shared-computer", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--odt-remove-msi", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument(
        "--odt-exclude-app",
        metavar="APP",
        action="append",
        dest="odt_exclude_apps",
        help=argparse.SUPPRESS,
    )
    parser.add_argument("--odt-include-visio", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--odt-include-project", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--odt-removal", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--odt-download", metavar="PATH", help=argparse.SUPPRESS)
    parser.add_argument("--odt-list-products", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--odt-list-presets", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--odt-list-channels", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--odt-list-languages", action="store_true", help=argparse.SUPPRESS)

    # Author Aliases
    parser.add_argument("--goobler", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--pupa", action="store_true", help=argparse.SUPPRESS)

    # OEM Configuration Presets
    oem_configs = parser.add_mutually_exclusive_group()
    oem_configs.add_argument(
        "--oem-config",
        metavar="NAME",
        choices=[
            "full-removal",
            "quick-repair",
            "full-repair",
            "proplus-x64",
            "proplus-x86",
            "proplus-visio-project",
            "business-x64",
            "office2019-x64",
            "office2021-x64",
            "office2024-x64",
            "multilang",
            "shared-computer",
            "interactive",
        ],
        help=argparse.SUPPRESS,
    )
    oem_configs.add_argument(
        "--c2r-remove",
        action="store_const",
        const="full-removal",
        dest="oem_config",
        help=argparse.SUPPRESS,
    )
    oem_configs.add_argument(
        "--c2r-repair-quick",
        action="store_const",
        const="quick-repair",
        dest="oem_config",
        help=argparse.SUPPRESS,
    )
    oem_configs.add_argument(
        "--c2r-repair-full",
        action="store_const",
        const="full-repair",
        dest="oem_config",
        help=argparse.SUPPRESS,
    )
    oem_configs.add_argument(
        "--c2r-proplus",
        action="store_const",
        const="proplus-x64",
        dest="oem_config",
        help=argparse.SUPPRESS,
    )
    oem_configs.add_argument(
        "--c2r-business",
        action="store_const",
        const="business-x64",
        dest="oem_config",
        help=argparse.SUPPRESS,
    )

    # OffScrub Legacy Compatibility
    parser.add_argument("--offscrub-all", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--offscrub-ose", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--offscrub-offline", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--offscrub-quiet", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--offscrub-test-rerun", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--offscrub-bypass", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--offscrub-fast-remove", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--offscrub-scan-components", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--offscrub-return-error", action="store_true", help=argparse.SUPPRESS)

    # Advanced Options
    parser.add_argument("--skip-preflight", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--skip-backup", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--skip-verification", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--schedule-reboot", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--no-schedule-delete", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--msiexec-args", metavar="ARGS", help=argparse.SUPPRESS)
    parser.add_argument("--c2r-args", metavar="ARGS", help=argparse.SUPPRESS)
    parser.add_argument("--odt-args", metavar="ARGS", help=argparse.SUPPRESS)

    return parser


def add_core_options(
    parser: argparse.ArgumentParser,
    *,
    include_passes: bool = True,
) -> argparse.ArgumentParser:
    """!
    @brief Add core operation options.
    @param parser The ArgumentParser to add arguments to.
    @param include_passes Whether to include the --passes option (default True).
        Set to False when the subparser already has a passes option.
    @returns The parser for chaining.
    """
    core = parser.add_argument_group("Core Options")
    core.add_argument(
        "--include",
        metavar="COMPONENTS",
        help="Additional suites/apps to include (visio,project,onenote).",
    )
    core.add_argument(
        "--force",
        "-f",
        action="store_true",
        help="Relax certain guardrails when safe.",
    )
    core.add_argument(
        "--allow-unsupported-windows",
        action="store_true",
        help="Permit execution on unsupported Windows versions.",
    )
    core.add_argument(
        "--dry-run",
        "-n",
        action="store_true",
        help="Simulate actions without modifying the system.",
    )
    core.add_argument(
        "--yes",
        "-y",
        action="store_true",
        help="Skip confirmation prompts (assume yes).",
    )
    core.add_argument(
        "--config",
        "-c",
        metavar="JSON",
        help="Load options from a JSON configuration file.",
    )
    if include_passes:
        core.add_argument(
            "--passes",
            type=int,
            default=None,
            metavar="N",
            help="Number of uninstall passes (default: 1).",
        )
    return parser


def add_repair_options(parser: argparse.ArgumentParser) -> argparse.ArgumentParser:
    """!
    @brief Add repair-specific options.
    @param parser The ArgumentParser to add arguments to.
    @returns The parser for chaining.
    """
    repair = parser.add_argument_group(
        "Repair Options",
        "Options for repairing Office installations.",
    )
    repair.add_argument(
        "--repair-odt",
        action="store_true",
        help="Repair using ODT/setup.exe configuration method.",
    )
    repair.add_argument(
        "--repair-c2r",
        action="store_true",
        help="Repair using OfficeClickToRun.exe directly.",
    )
    repair.add_argument(
        "--repair-culture",
        metavar="LANG",
        default="en-us",
        help="Language/culture code for repair (default: en-us).",
    )
    repair.add_argument(
        "--repair-platform",
        choices=["x86", "x64"],
        metavar="ARCH",
        help="Architecture for repair (auto-detected if not specified).",
    )
    repair.add_argument(
        "--repair-visible",
        action="store_true",
        help="Show repair UI instead of running silently.",
    )
    repair.add_argument(
        "--repair-timeout",
        type=int,
        default=3600,
        metavar="SEC",
        help="Timeout for repair operations in seconds (default: 3600).",
    )
    repair.add_argument(
        "--repair-all-products",
        action="store_true",
        help="Repair all detected Office products (for auto-repair).",
    )
    repair.add_argument(
        "--repair-preset",
        metavar="NAME",
        help="Use a specific repair preset (quick-repair, full-repair, etc.).",
    )
    return parser


def add_uninstall_options(parser: argparse.ArgumentParser) -> argparse.ArgumentParser:
    """!
    @brief Add uninstall method options.
    @param parser The ArgumentParser to add arguments to.
    @returns The parser for chaining.
    """
    uninstall = parser.add_argument_group("Uninstall Method Options")
    uninstall.add_argument(
        "--uninstall-method",
        choices=["auto", "msi", "c2r", "odt", "offscrub"],
        default=None,
        metavar="METHOD",
        help="Preferred uninstall method: auto, msi, c2r, odt, offscrub.",
    )
    uninstall.add_argument(
        "--msi-only",
        action="store_const",
        const="msi",
        dest="uninstall_method",
        help="Only uninstall MSI-based Office products.",
    )
    uninstall.add_argument(
        "--c2r-only",
        action="store_const",
        const="c2r",
        dest="uninstall_method",
        help="Only uninstall Click-to-Run Office products.",
    )
    uninstall.add_argument(
        "--use-odt",
        action="store_const",
        const="odt",
        dest="uninstall_method",
        help="Use Office Deployment Tool (setup.exe) for uninstall.",
    )
    uninstall.add_argument(
        "--force-app-shutdown",
        action="store_true",
        help="Force close running Office applications before uninstall.",
    )
    uninstall.add_argument(
        "--no-force-app-shutdown",
        action="store_true",
        help="Prompt user to close apps instead of forcing shutdown.",
    )
    uninstall.add_argument(
        "--product-code",
        metavar="GUID",
        action="append",
        dest="product_codes",
        help="Specific MSI product code(s) to uninstall (repeatable).",
    )
    uninstall.add_argument(
        "--release-id",
        metavar="ID",
        action="append",
        dest="release_ids",
        help="Specific C2R release ID(s) to uninstall (repeatable).",
    )
    return parser


def add_scrub_options(parser: argparse.ArgumentParser) -> argparse.ArgumentParser:
    """!
    @brief Add scrubbing options.
    @param parser The ArgumentParser to add arguments to.
    @returns The parser for chaining.
    """
    scrub = parser.add_argument_group("Scrubbing Options")
    scrub.add_argument(
        "--scrub-level",
        choices=["minimal", "standard", "aggressive", "nuclear"],
        default=None,
        metavar="LEVEL",
        help="Scrub intensity level (minimal|standard|aggressive|nuclear).",
    )
    scrub.add_argument(
        "--max-passes",
        type=int,
        default=None,
        metavar="N",
        help="Maximum uninstall/re-detect passes (alias for --passes).",
    )
    scrub.add_argument(
        "--skip-uninstall",
        action="store_true",
        help="Skip uninstall passes; only run cleanup/registry scrubbing.",
    )
    scrub.add_argument(
        "--skip-processes",
        action="store_true",
        help="Skip terminating Office processes before uninstall.",
    )
    scrub.add_argument(
        "--skip-services",
        action="store_true",
        help="Skip stopping Office services before uninstall.",
    )
    scrub.add_argument(
        "--skip-tasks",
        action="store_true",
        help="Skip removing scheduled tasks.",
    )
    scrub.add_argument(
        "--skip-registry",
        action="store_true",
        help="Skip registry cleanup after uninstall.",
    )
    scrub.add_argument(
        "--skip-filesystem",
        action="store_true",
        help="Skip filesystem cleanup after uninstall.",
    )
    scrub.add_argument(
        "--registry-only",
        action="store_true",
        help="Only perform registry cleanup.",
    )
    scrub.add_argument(
        "--clean-msocache",
        action="store_true",
        help="Also remove MSOCache installation files.",
    )
    scrub.add_argument(
        "--clean-appx",
        action="store_true",
        help="Also remove Office AppX/MSIX packages.",
    )
    scrub.add_argument(
        "--clean-wi-metadata",
        action="store_true",
        help="Clean orphaned Windows Installer metadata.",
    )
    return parser


def add_license_options(parser: argparse.ArgumentParser) -> argparse.ArgumentParser:
    """!
    @brief Add license and activation options.
    @param parser The ArgumentParser to add arguments to.
    @returns The parser for chaining.
    """
    license_opts = parser.add_argument_group("License & Activation Options")
    restore_point = license_opts.add_mutually_exclusive_group()
    restore_point.add_argument(
        "--restore-point",
        "--create-restore-point",
        action="store_true",
        dest="create_restore_point",
        help="Create a system restore point before scrubbing.",
    )
    restore_point.add_argument(
        "--no-restore-point",
        action="store_true",
        help="Skip creating a system restore point.",
    )
    license_opts.add_argument(
        "--no-license",
        action="store_true",
        help="Skip license cleanup steps.",
    )
    license_opts.add_argument(
        "--keep-license",
        action="store_true",
        help="Preserve Office licenses (alias of --no-license).",
    )
    license_opts.add_argument(
        "--clean-spp",
        action="store_true",
        help="Clean Software Protection Platform (SPP) Office tokens.",
    )
    license_opts.add_argument(
        "--clean-ospp",
        action="store_true",
        help="Clean Office Software Protection Platform (OSPP) tokens.",
    )
    license_opts.add_argument(
        "--clean-vnext",
        action="store_true",
        help="Clean vNext/device-based licensing cache.",
    )
    license_opts.add_argument(
        "--clean-all-licenses",
        action="store_true",
        help="Aggressively clean all license artifacts.",
    )
    return parser


def add_data_options(parser: argparse.ArgumentParser) -> argparse.ArgumentParser:
    """!
    @brief Add user data options.
    @param parser The ArgumentParser to add arguments to.
    @returns The parser for chaining.
    """
    data = parser.add_argument_group("User Data Options")
    data.add_argument(
        "--keep-templates",
        action="store_true",
        help="Preserve user templates like normal.dotm.",
    )
    data.add_argument(
        "--keep-user-settings",
        action="store_true",
        help="Preserve user Office settings and customizations.",
    )
    data.add_argument(
        "--delete-user-settings",
        action="store_true",
        help="Remove user Office settings and customizations.",
    )
    data.add_argument(
        "--keep-outlook-data",
        action="store_true",
        help="Preserve Outlook OST/PST files and profiles.",
    )
    data.add_argument(
        "--keep-outlook-signatures",
        action="store_true",
        help="Preserve Outlook email signatures when deleting Outlook data.",
    )
    data.add_argument(
        "--clean-shortcuts",
        action="store_true",
        help="Remove Office shortcuts from Start Menu and Desktop.",
    )
    data.add_argument(
        "--skip-shortcut-detection",
        action="store_true",
        help="Skip detecting and cleaning orphaned shortcuts.",
    )
    return parser


def add_registry_cleanup_options(parser: argparse.ArgumentParser) -> argparse.ArgumentParser:
    """!
    @brief Add registry cleanup options.
    @param parser The ArgumentParser to add arguments to.
    @returns The parser for chaining.
    """
    registry = parser.add_argument_group("Registry Cleanup Options")
    registry.add_argument(
        "--clean-addin-registry",
        action="store_true",
        help="Clean Office add-in registry entries.",
    )
    registry.add_argument(
        "--clean-com-registry",
        action="store_true",
        help="Clean orphaned COM/ActiveX registrations.",
    )
    registry.add_argument(
        "--clean-shell-extensions",
        action="store_true",
        help="Clean orphaned shell extensions.",
    )
    registry.add_argument(
        "--clean-typelibs",
        action="store_true",
        help="Clean orphaned type libraries.",
    )
    registry.add_argument(
        "--clean-protocol-handlers",
        action="store_true",
        help="Clean Office protocol handlers.",
    )
    registry.add_argument(
        "--remove-vba",
        action="store_true",
        help="Remove VBA-only package and registry entries.",
    )
    return parser


def add_output_options(
    parser: argparse.ArgumentParser,
    *,
    include_timeout: bool = True,
    include_quiet: bool = True,
) -> argparse.ArgumentParser:
    """!
    @brief Add output and logging options.
    @param parser The ArgumentParser to add arguments to.
    @param include_timeout Whether to include the --timeout option (default True).
        Set to False when the subparser already has a timeout option.
    @param include_quiet Whether to include the --quiet option (default True).
        Set to False when the subparser already defines its own quiet option.
    @returns The parser for chaining.
    """
    output = parser.add_argument_group("Output & Logging Options")
    output.add_argument(
        "--plan",
        metavar="OUT",
        help="Write the computed action plan to a JSON file.",
    )
    output.add_argument(
        "--logdir",
        metavar="DIR",
        help="Directory for human/JSONL log output.",
    )
    output.add_argument(
        "--backup",
        metavar="DIR",
        help="Destination for registry/file backups.",
    )
    if include_timeout:
        output.add_argument(
            "--timeout",
            metavar="SEC",
            type=int,
            help="Per-step timeout in seconds.",
        )
    if include_quiet:
        output.add_argument(
            "--quiet",
            "-q",
            action="store_true",
            help="Minimal console output (errors only).",
        )
    output.add_argument(
        "--json",
        action="store_true",
        help="Mirror structured events to stdout.",
    )
    output.add_argument(
        "--verbose",
        "-v",
        action="count",
        default=0,
        help="Increase output verbosity (-v, -vv, -vvv).",
    )
    return parser


def add_tui_options(parser: argparse.ArgumentParser) -> argparse.ArgumentParser:
    """!
    @brief Add TUI options.
    @param parser The ArgumentParser to add arguments to.
    @returns The parser for chaining.
    """
    tui = parser.add_argument_group("TUI Options")
    tui.add_argument(
        "--tui",
        action="store_true",
        help="Force the interactive text UI mode.",
    )
    tui.add_argument(
        "--no-color",
        action="store_true",
        help="Disable ANSI color codes.",
    )
    tui.add_argument(
        "--tui-compact",
        action="store_true",
        help="Use a compact TUI layout for small consoles.",
    )
    tui.add_argument(
        "--tui-refresh",
        metavar="MS",
        type=int,
        help="Refresh interval for the TUI renderer in milliseconds.",
    )
    tui.add_argument(
        "--limited-user",
        action="store_true",
        help="Run detection under a limited user token when possible.",
    )
    return parser


def add_retry_options(parser: argparse.ArgumentParser) -> argparse.ArgumentParser:
    """!
    @brief Add retry and resilience options.
    @param parser The ArgumentParser to add arguments to.
    @returns The parser for chaining.
    """
    retry = parser.add_argument_group("Retry & Resilience Options")
    retry.add_argument(
        "--retries",
        type=int,
        default=None,
        metavar="N",
        help="Number of retry attempts per step (default: 4).",
    )
    retry.add_argument(
        "--retry-delay",
        type=int,
        default=3,
        metavar="SEC",
        help="Base delay between retries in seconds (default: 3).",
    )
    retry.add_argument(
        "--retry-delay-max",
        type=int,
        default=30,
        metavar="SEC",
        help="Maximum delay between retries in seconds (default: 30).",
    )
    retry.add_argument(
        "--no-reboot",
        action="store_true",
        help="Suppress reboot recommendations.",
    )
    retry.add_argument(
        "--offline",
        action="store_true",
        help="Run in offline mode (no network access for downloads).",
    )
    return parser


def add_odt_build_options(parser: argparse.ArgumentParser) -> argparse.ArgumentParser:
    """!
    @brief Add ODT build options.
    @param parser The ArgumentParser to add arguments to.
    @returns The parser for chaining.
    """
    odt = parser.add_argument_group(
        "ODT Build Options",
        "Generate Office Deployment Tool XML configurations.",
    )
    odt.add_argument(
        "--odt-build",
        action="store_true",
        help="Generate an ODT XML configuration file.",
    )
    odt.add_argument(
        "--odt-install",
        action="store_true",
        help="Install Office using ODT with the specified configuration.",
    )
    odt.add_argument(
        "--odt-preset",
        metavar="NAME",
        help="Use a predefined ODT installation preset.",
    )
    odt.add_argument(
        "--odt-product",
        metavar="ID",
        action="append",
        dest="odt_products",
        help="Product ID to include in ODT config (repeatable).",
    )
    odt.add_argument(
        "--odt-language",
        metavar="CODE",
        action="append",
        dest="odt_languages",
        help="Language code for ODT config (repeatable, default: en-us).",
    )
    odt.add_argument(
        "--odt-arch",
        choices=["32", "64"],
        default="64",
        metavar="BITS",
        help="Architecture for ODT config (default: 64).",
    )
    odt.add_argument(
        "--odt-channel",
        metavar="CHANNEL",
        help="Update channel for ODT config.",
    )
    odt.add_argument(
        "--odt-output",
        metavar="FILE",
        help="Output path for generated ODT XML.",
    )
    odt.add_argument(
        "--odt-shared-computer",
        action="store_true",
        help="Enable shared computer licensing.",
    )
    odt.add_argument(
        "--odt-remove-msi",
        action="store_true",
        help="Include RemoveMSI element.",
    )
    odt.add_argument(
        "--odt-exclude-app",
        metavar="APP",
        action="append",
        dest="odt_exclude_apps",
        help="App to exclude from installation (repeatable).",
    )
    odt.add_argument(
        "--odt-include-visio",
        action="store_true",
        help="Include Visio Professional.",
    )
    odt.add_argument(
        "--odt-include-project",
        action="store_true",
        help="Include Project Professional.",
    )
    odt.add_argument(
        "--odt-removal",
        action="store_true",
        help="Generate a removal XML instead of installation XML.",
    )
    odt.add_argument(
        "--odt-download",
        metavar="PATH",
        help="Generate a download XML with the specified local path.",
    )
    odt.add_argument(
        "--odt-list-products",
        action="store_true",
        help="List all available ODT product IDs and exit.",
    )
    odt.add_argument(
        "--odt-list-presets",
        action="store_true",
        help="List all available ODT installation presets and exit.",
    )
    odt.add_argument(
        "--odt-list-channels",
        action="store_true",
        help="List all available ODT update channels and exit.",
    )
    odt.add_argument(
        "--odt-list-languages",
        action="store_true",
        help="List all supported language codes and exit.",
    )
    return parser


def add_author_aliases(parser: argparse.ArgumentParser) -> argparse.ArgumentParser:
    """!
    @brief Add author quick install aliases.
    @param parser The ArgumentParser to add arguments to.
    @returns The parser for chaining.
    """
    author = parser.add_argument_group(
        "Quick Install Aliases",
        "Author-defined shortcuts for common Office installations.",
    )
    author.add_argument(
        "--goobler",
        action="store_true",
        help="Install LTSC 2024 + Visio + Project (clean) with pt-pt and en-us.",
    )
    author.add_argument(
        "--pupa",
        action="store_true",
        help="Install LTSC 2024 ProPlus only (clean) with pt-pt and en-us.",
    )
    return parser


def add_oem_config_options(parser: argparse.ArgumentParser) -> argparse.ArgumentParser:
    """!
    @brief Add OEM configuration preset options.
    @param parser The ArgumentParser to add arguments to.
    @returns The parser for chaining.
    """
    oem = parser.add_argument_group("OEM Configuration Presets")
    oem_configs = oem.add_mutually_exclusive_group()
    oem_configs.add_argument(
        "--oem-config",
        metavar="NAME",
        choices=[
            "full-removal",
            "quick-repair",
            "full-repair",
            "proplus-x64",
            "proplus-x86",
            "proplus-visio-project",
            "business-x64",
            "office2019-x64",
            "office2021-x64",
            "office2024-x64",
            "multilang",
            "shared-computer",
            "interactive",
        ],
        help="Use bundled OEM configuration preset.",
    )
    oem_configs.add_argument(
        "--c2r-remove",
        action="store_const",
        const="full-removal",
        dest="oem_config",
        help="Remove all Office C2R products.",
    )
    oem_configs.add_argument(
        "--c2r-repair-quick",
        action="store_const",
        const="quick-repair",
        dest="oem_config",
        help="Quick repair Office C2R.",
    )
    oem_configs.add_argument(
        "--c2r-repair-full",
        action="store_const",
        const="full-repair",
        dest="oem_config",
        help="Full online repair Office C2R.",
    )
    oem_configs.add_argument(
        "--c2r-proplus",
        action="store_const",
        const="proplus-x64",
        dest="oem_config",
        help="Repair Office 365 ProPlus x64.",
    )
    oem_configs.add_argument(
        "--c2r-business",
        action="store_const",
        const="business-x64",
        dest="oem_config",
        help="Repair Microsoft 365 Business x64.",
    )
    return parser


def add_offscrub_options(parser: argparse.ArgumentParser) -> argparse.ArgumentParser:
    """!
    @brief Add OffScrub legacy compatibility options.
    @param parser The ArgumentParser to add arguments to.
    @returns The parser for chaining.
    """
    offscrub = parser.add_argument_group(
        "OffScrub Legacy Compatibility",
        "Flags for compatibility with legacy OffScrub VBS scripts.",
    )
    offscrub.add_argument(
        "--offscrub-all",
        action="store_true",
        help="OffScrub /ALL: Remove all detected Office products.",
    )
    offscrub.add_argument(
        "--offscrub-ose",
        action="store_true",
        help="OffScrub /OSE: Fix OSE service configuration.",
    )
    offscrub.add_argument(
        "--offscrub-offline",
        action="store_true",
        help="OffScrub /OFFLINE: Mark C2R config as offline mode.",
    )
    offscrub.add_argument(
        "--offscrub-quiet",
        action="store_true",
        help="OffScrub /QUIET: Reduce human log verbosity.",
    )
    offscrub.add_argument(
        "--offscrub-test-rerun",
        action="store_true",
        help="OffScrub /TR: Run uninstall passes twice.",
    )
    offscrub.add_argument(
        "--offscrub-bypass",
        action="store_true",
        help="OffScrub /BYPASS: Bypass certain safety checks.",
    )
    offscrub.add_argument(
        "--offscrub-fast-remove",
        action="store_true",
        help="OffScrub /FASTREMOVE: Skip verification probes.",
    )
    offscrub.add_argument(
        "--offscrub-scan-components",
        action="store_true",
        help="OffScrub /SCANCOMPONENTS: Scan Windows Installer.",
    )
    offscrub.add_argument(
        "--offscrub-return-error",
        action="store_true",
        help="OffScrub /RETERRORSUCCESS: Return error codes on partial.",
    )
    return parser


def add_advanced_options(parser: argparse.ArgumentParser) -> argparse.ArgumentParser:
    """!
    @brief Add advanced options.
    @param parser The ArgumentParser to add arguments to.
    @returns The parser for chaining.
    """
    advanced = parser.add_argument_group("Advanced Options")
    advanced.add_argument(
        "--skip-preflight",
        action="store_true",
        help="Skip preflight safety checks.",
    )
    advanced.add_argument(
        "--skip-backup",
        action="store_true",
        help="Skip creating registry and file backups.",
    )
    advanced.add_argument(
        "--skip-verification",
        action="store_true",
        help="Skip verification probes after uninstall.",
    )
    advanced.add_argument(
        "--schedule-reboot",
        action="store_true",
        help="Schedule a reboot after completion if recommended.",
    )
    advanced.add_argument(
        "--no-schedule-delete",
        action="store_true",
        help="Don't use MoveFileEx for locked file deletion.",
    )
    advanced.add_argument(
        "--msiexec-args",
        metavar="ARGS",
        help="Additional arguments to pass to msiexec.",
    )
    advanced.add_argument(
        "--c2r-args",
        metavar="ARGS",
        help="Additional arguments to pass to OfficeC2RClient.exe.",
    )
    advanced.add_argument(
        "--odt-args",
        metavar="ARGS",
        help="Additional arguments to pass to ODT setup.exe.",
    )
    return parser


def build_arg_parser(version_info: dict[str, str] | None = None) -> argparse.ArgumentParser:
    """!
    @brief Create the top-level argument parser with subcommands and all CLI options.
    @param version_info Optional version metadata dict with 'version' and 'build' keys.
    @returns Configured ArgumentParser instance.
    @details Supports three operation modes via subcommands:
    - install: Deploy Office via ODT presets or custom configurations
    - repair: Fix broken Office installations
    - remove: Uninstall Office and clean up residual artifacts

    Legacy flags (--auto-all, --auto-repair, etc.) are preserved for backward compatibility.
    """
    parser = argparse.ArgumentParser(
        prog=PROGRAM_NAME,
        add_help=False,  # We'll add custom help action
        description=PROGRAM_DESCRIPTION,
        epilog=EPILOG_TEXT,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    # Add custom help action that pauses before exit
    parser.add_argument(
        "-h",
        "--help",
        action=HelpActionWithPause,
        help="show this help message and exit",
    )

    if version_info:
        parser.add_argument(
            "-V",
            "--version",
            action=VersionActionWithPause,
            version=f"{version_info.get('version', '0.0.0')} ({version_info.get('build', 'dev')})",
        )

    # ---------------------------------------------------------------------------
    # Subcommands: install, repair, remove
    # ---------------------------------------------------------------------------
    subparsers = parser.add_subparsers(
        dest="command",
        title="operation modes",
        description="Choose an operation mode (or use legacy flags for backward compatibility)",
        metavar="<command>",
    )

    # ----- INSTALL subcommand -----
    install_parser = subparsers.add_parser(
        "install",
        help="Deploy Office via ODT presets or custom configurations",
        description="Install Microsoft Office using the Office Deployment Tool (ODT).",
        epilog=INSTALL_EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
        add_help=False,
    )
    install_parser.add_argument("-h", "--help", action=HelpActionWithPause)
    install_parser.set_defaults(show_help=install_parser)
    add_install_subcommand_options(install_parser)
    add_core_options(install_parser)
    add_output_options(install_parser)
    add_tui_options(install_parser)
    add_retry_options(install_parser)
    add_advanced_options(install_parser)

    # ----- REPAIR subcommand -----
    repair_parser = subparsers.add_parser(
        "repair",
        help="Fix broken Office installations (quick or full repair)",
        description="Repair Microsoft Office installations using various methods.",
        epilog=REPAIR_EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
        add_help=False,
    )
    repair_parser.add_argument("-h", "--help", action=HelpActionWithPause)
    repair_parser.set_defaults(show_help=repair_parser)
    add_repair_subcommand_options(repair_parser)
    add_core_options(repair_parser)
    add_output_options(repair_parser, include_timeout=False)  # Timeout already in repair opts
    add_tui_options(repair_parser)
    add_retry_options(repair_parser)
    add_advanced_options(repair_parser)

    # ----- REMOVE subcommand -----
    remove_parser = subparsers.add_parser(
        "remove",
        help="Uninstall Office and clean up residual artifacts",
        description="Remove Microsoft Office installations and scrub leftover artifacts.",
        epilog=REMOVE_EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
        add_help=False,
    )
    remove_parser.add_argument("-h", "--help", action=HelpActionWithPause)
    remove_parser.set_defaults(show_help=remove_parser)
    add_remove_subcommand_options(remove_parser)
    add_core_options(remove_parser, include_passes=False)  # Passes already in remove opts
    add_scrub_options(remove_parser)
    add_license_options(remove_parser)
    add_data_options(remove_parser)
    add_registry_cleanup_options(remove_parser)
    add_output_options(remove_parser)
    add_tui_options(remove_parser)
    add_retry_options(remove_parser)
    add_offscrub_options(remove_parser)
    add_advanced_options(remove_parser)

    # ----- DIAGNOSE subcommand -----
    diagnose_parser = subparsers.add_parser(
        "diagnose",
        help="Detect and report Office installations without making changes",
        description="Scan the system for Office installations and generate a diagnostic report.",
        epilog=DIAGNOSE_EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
        add_help=False,
    )
    diagnose_parser.add_argument("-h", "--help", action=HelpActionWithPause)
    diagnose_parser.set_defaults(show_help=diagnose_parser)
    add_diagnose_subcommand_options(diagnose_parser)
    add_output_options(diagnose_parser)
    add_tui_options(diagnose_parser)

    # ----- ODT subcommand -----
    odt_parser = subparsers.add_parser(
        "odt",
        help="Build and manage Office Deployment Tool XML configurations",
        description="Generate ODT configuration files for Office installation/removal.",
        epilog=ODT_EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
        add_help=False,
    )
    odt_parser.add_argument("-h", "--help", action=HelpActionWithPause)
    odt_parser.set_defaults(show_help=odt_parser)
    add_odt_subcommand_options(odt_parser)
    add_output_options(odt_parser)
    add_tui_options(odt_parser)

    # ----- OFFSCRUB subcommand -----
    offscrub_parser = subparsers.add_parser(
        "offscrub",
        help="OffScrub-style deep removal of Office installations",
        description="Perform thorough Office removal using OffScrub techniques.",
        epilog=OFFSCRUB_EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
        add_help=False,
    )
    offscrub_parser.add_argument("-h", "--help", action=HelpActionWithPause)
    offscrub_parser.set_defaults(show_help=offscrub_parser)
    add_offscrub_subcommand_options(offscrub_parser)
    add_output_options(offscrub_parser, include_quiet=False)  # offscrub defines own -Q/--quiet
    add_tui_options(offscrub_parser)
    add_advanced_options(offscrub_parser)

    # ----- C2R subcommand -----
    c2r_parser = subparsers.add_parser(
        "c2r",
        help="Direct Click-to-Run operations passthrough",
        description="Execute Click-to-Run operations directly via OfficeClickToRun.exe.",
        epilog=C2R_EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
        add_help=False,
    )
    c2r_parser.add_argument("-h", "--help", action=HelpActionWithPause)
    c2r_parser.set_defaults(show_help=c2r_parser)
    add_c2r_subcommand_options(c2r_parser)
    add_output_options(c2r_parser, include_timeout=False)  # c2r defines own --timeout
    add_tui_options(c2r_parser)

    # ----- LICENSE subcommand -----
    license_parser = subparsers.add_parser(
        "license",
        help="Manage Office licensing and activation",
        description="View, clean, and manage Office licensing information.",
        epilog=LICENSE_EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
        add_help=False,
    )
    license_parser.add_argument("-h", "--help", action=HelpActionWithPause)
    license_parser.set_defaults(show_help=license_parser)
    add_license_subcommand_options(license_parser)
    add_output_options(license_parser)
    add_tui_options(license_parser)

    # ----- CONFIG subcommand -----
    config_parser = subparsers.add_parser(
        "config",
        help="Generate and manage configuration files",
        description="Generate Office Janitor configuration files interactively or from templates.",
        epilog=CONFIG_EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
        add_help=False,
    )
    config_parser.add_argument("-h", "--help", action=HelpActionWithPause)
    config_parser.set_defaults(show_help=config_parser)
    add_config_subcommand_options(config_parser)
    add_output_options(config_parser)
    add_tui_options(config_parser)

    # ---------------------------------------------------------------------------
    # Global options (shown in main help)
    # ---------------------------------------------------------------------------
    # Truly global options that apply to all modes/subcommands
    add_mode_arguments(parser)  # Legacy mode flags (--auto-all, --diagnose, etc.)
    add_global_options(parser)  # Truly global: --dry-run, --yes, --verbose, etc.

    # ---------------------------------------------------------------------------
    # Legacy options (hidden from main help, for backward compatibility)
    # ---------------------------------------------------------------------------
    # These options work without subcommands for backward compatibility.
    # They are suppressed from help to keep the main help clean - users should
    # use subcommand help for detailed options.
    add_legacy_options(parser)

    return parser


# ---------------------------------------------------------------------------
# Help Text Generation Utilities
# ---------------------------------------------------------------------------


def format_repair_help() -> str:
    """!
    @brief Generate detailed help text for repair operations.
    @returns Formatted help string for repair operations.
    """
    return """\
OFFICE REPAIR OPTIONS
=====================

Office Janitor provides multiple repair mechanisms for Office installations:

AUTO-REPAIR MODE (--auto-repair)
--------------------------------
Automatically detects all Office installations and repairs them using the
most appropriate method. Supports both MSI and Click-to-Run installations.

  office-janitor --auto-repair              # Repair all detected Office
  office-janitor --auto-repair --dry-run    # Preview repair operations
  office-janitor --auto-repair --force      # Skip confirmations

QUICK REPAIR (--repair quick)
-----------------------------
Local repair that doesn't require internet. Fixes corrupted files and
common issues. Typically completes in 5-15 minutes.

  office-janitor --repair quick
  office-janitor --repair quick --repair-culture de-de

FULL ONLINE REPAIR (--repair full)
----------------------------------
Downloads and reinstalls Office components from Microsoft CDN.
More thorough but requires internet and takes 30-60 minutes.
WARNING: May reinstall previously excluded applications.

  office-janitor --repair full
  office-janitor --repair full --repair-visible

ODT REPAIR (--repair-odt)
-------------------------
Uses Office Deployment Tool configuration to repair/reconfigure Office.
Useful for enterprise deployments and custom configurations.

  office-janitor --repair-odt --repair-preset proplus-x64
  office-janitor --repair-odt --repair-config custom.xml

C2R REPAIR (--repair-c2r)
-------------------------
Direct repair using OfficeClickToRun.exe. Provides granular control
over repair parameters.

  office-janitor --repair-c2r --repair-culture en-us
  office-janitor --repair-c2r --repair-platform x64
"""


def format_quick_reference() -> str:
    """!
    @brief Generate a quick reference card for common operations.
    @returns Formatted quick reference string.
    """
    return """\
OFFICE JANITOR QUICK REFERENCE
==============================

REPAIR:
  --auto-repair              Auto-detect and repair all Office installations
  --repair quick             Quick local repair (5-15 min, no internet)
  --repair full              Full online repair (30-60 min, needs internet)
  --repair-odt               Repair using ODT configuration
  --repair-c2r               Repair using C2R client directly

REMOVE:
  --auto-all                 Full detection and removal of all Office
  --target VERSION           Target specific version (2013, 2016, 2019, etc.)
  --c2r-remove               Remove Click-to-Run Office only

INSTALL:
  --odt-install --odt-preset NAME    Install using predefined preset
  --goobler                          LTSC 2024 + Visio + Project (author alias)
  --pupa                             LTSC 2024 ProPlus (author alias)

DIAGNOSE:
  --diagnose                 Generate inventory without changes
  --plan output.json         Save action plan to file

COMMON OPTIONS:
  --dry-run, -n              Preview without making changes
  --force, -f                Skip confirmation prompts
  --verbose, -v              Increase output verbosity
  --quiet, -q                Minimal output (errors only)
"""


__all__ = [
    "PROGRAM_NAME",
    "PROGRAM_DESCRIPTION",
    "EPILOG_TEXT",
    "INSTALL_EPILOG",
    "REPAIR_EPILOG",
    "REMOVE_EPILOG",
    "DIAGNOSE_EPILOG",
    "build_arg_parser",
    "format_repair_help",
    "format_quick_reference",
    "add_mode_arguments",
    "add_core_options",
    "add_repair_options",
    "add_uninstall_options",
    "add_scrub_options",
    "add_license_options",
    "add_data_options",
    "add_registry_cleanup_options",
    "add_output_options",
    "add_tui_options",
    "add_retry_options",
    "add_odt_build_options",
    "add_author_aliases",
    "add_oem_config_options",
    "add_offscrub_options",
    "add_advanced_options",
]
