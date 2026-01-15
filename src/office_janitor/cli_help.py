"""!
@brief CLI help text and argument definitions for Office Janitor.
@details This module centralizes all CLI argument parsing and help text generation,
enabling a streamlined help system and consistent command-line interface.
"""

from __future__ import annotations

import argparse
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    pass

# ---------------------------------------------------------------------------
# Version & Program Info
# ---------------------------------------------------------------------------

PROGRAM_NAME = "office-janitor"
PROGRAM_DESCRIPTION = "Detect, repair, uninstall, and scrub Microsoft Office installations."

# ---------------------------------------------------------------------------
# Help Text Constants
# ---------------------------------------------------------------------------

EPILOG_TEXT = """\
================================================================================
                           QUICK REFERENCE GUIDE
================================================================================

REPAIR OPERATIONS:
  office-janitor --auto-repair              Auto-detect and repair all Office
  office-janitor --repair quick             Quick local repair (5-15 min)
  office-janitor --repair full              Full online repair (30-60 min)
  office-janitor --repair-odt               Repair via ODT configuration
  office-janitor --repair-c2r               Repair via C2R client directly

INSTALLATION PRESETS (--odt-install --odt-preset NAME):
  365-proplus-x64              Microsoft 365 Apps for enterprise (64-bit)
  365-business-x64             Microsoft 365 Apps for business (64-bit)
  office2024-x64               Office LTSC 2024 Professional Plus (64-bit)
  office2021-x64               Office LTSC 2021 Professional Plus (64-bit)
  ltsc2024-full-x64            Office 2024 + Visio + Project (64-bit)
  ltsc2024-full-x64-clean      Office 2024 + Visio + Project (no bloat)

REMOVAL OPERATIONS:
  office-janitor --auto-all                 Full detection and removal
  office-janitor --auto-all --dry-run       Preview removal (safe)
  office-janitor --target 2019              Target specific Office version
  office-janitor --c2r-remove               Remove C2R Office only

DIAGNOSTICS:
  office-janitor --diagnose                 Emit inventory without changes
  office-janitor --diagnose --plan out.json Save detailed plan to file

================================================================================

SUPPORTED LANGUAGES: en-us, de-de, fr-fr, es-es, pt-br, pt-pt, it-it, ja-jp,
  ko-kr, zh-cn, zh-tw, ru-ru, pl-pl, nl-nl, ar-sa, he-il, tr-tr

Use --odt-list-languages for the complete list of 60+ supported language codes.
Use --odt-list-products for all available Office product IDs.
Use --odt-list-channels for update channel options.
Use --odt-list-presets for all installation presets.

Full documentation: https://github.com/supermarsx/office-janitor
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


def add_core_options(parser: argparse.ArgumentParser) -> argparse.ArgumentParser:
    """!
    @brief Add core operation options.
    @param parser The ArgumentParser to add arguments to.
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


def add_output_options(parser: argparse.ArgumentParser) -> argparse.ArgumentParser:
    """!
    @brief Add output and logging options.
    @param parser The ArgumentParser to add arguments to.
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
    output.add_argument(
        "--timeout",
        metavar="SEC",
        type=int,
        help="Per-step timeout in seconds.",
    )
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
    @brief Create the top-level argument parser with all CLI options.
    @param version_info Optional version metadata dict with 'version' and 'build' keys.
    @returns Configured ArgumentParser instance.
    """
    parser = argparse.ArgumentParser(
        prog=PROGRAM_NAME,
        add_help=True,
        description=PROGRAM_DESCRIPTION,
        epilog=EPILOG_TEXT,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    if version_info:
        parser.add_argument(
            "-V",
            "--version",
            action="version",
            version=f"{version_info.get('version', '0.0.0')} ({version_info.get('build', 'dev')})",
        )

    # Add all argument groups in logical order
    add_mode_arguments(parser)
    add_core_options(parser)
    add_repair_options(parser)
    add_uninstall_options(parser)
    add_scrub_options(parser)
    add_license_options(parser)
    add_data_options(parser)
    add_registry_cleanup_options(parser)
    add_output_options(parser)
    add_tui_options(parser)
    add_retry_options(parser)
    add_odt_build_options(parser)
    add_author_aliases(parser)
    add_oem_config_options(parser)
    add_offscrub_options(parser)
    add_advanced_options(parser)

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
