"""!
@brief Primary entry point for the Office Janitor CLI.
@details This module bootstraps the runtime environment described in
:mod:`spec.md`: argument parsing for detection and scrubbing modes, validating
administrative elevation, enabling Windows VT mode when available, and invoking
logging setup so future sub-systems can emit structured telemetry.
"""

from __future__ import annotations

import argparse
import ctypes
import datetime
import json
import logging
import os
import pathlib
import platform
import signal
import sys
import threading
import time
from collections.abc import Iterable, Mapping

from . import (
    confirm,
    constants,
    detect,
    elevation,
    exec_utils,
    fs_tools,
    logging_ext,
    odt_build,
    processes,
    repair,
    safety,
    scrub,
    spinner,
    tui,
    ui,
    version,
)
from . import (
    plan as plan_module,
)
from .app_state import AppState, new_event_queue

# ---------------------------------------------------------------------------
# Progress logging utilities
# ---------------------------------------------------------------------------

_MAIN_START_TIME: float | None = None
_PROGRESS_LOCK = threading.Lock()  # Thread-safe progress output
_PENDING_LINE_OWNER: int | None = None  # Thread ID that owns the current incomplete line


def _get_elapsed_secs() -> float:
    """Return seconds since main() started, or 0 if not started."""
    if _MAIN_START_TIME is None:
        return 0.0
    return time.perf_counter() - _MAIN_START_TIME


def _progress(message: str, *, newline: bool = True, indent: int = 0) -> None:
    """Print a progress message with dmesg-style timestamp."""
    global _PENDING_LINE_OWNER
    with _PROGRESS_LOCK:
        # Clear spinner before output
        spinner.pause_for_output()
        try:
            # If another thread left an incomplete line, finish it first
            if _PENDING_LINE_OWNER is not None and _PENDING_LINE_OWNER != threading.get_ident():
                print(flush=True)  # Force newline
                _PENDING_LINE_OWNER = None
                spinner.clear_incomplete_line()

            timestamp = f"[{_get_elapsed_secs():12.6f}]"
            prefix = "  " * indent
            if newline:
                print(f"{timestamp} {prefix}{message}", flush=True)
                _PENDING_LINE_OWNER = None
                spinner.resume_after_output()
            else:
                print(f"{timestamp} {prefix}{message}", end="", flush=True)
                _PENDING_LINE_OWNER = threading.get_ident()
                spinner.mark_incomplete_line()
                spinner.resume_after_output()
        except Exception:
            spinner.resume_after_output()
            raise


def _progress_ok(extra: str = "") -> None:
    """Print OK status in Linux init style [  OK  ]."""
    global _PENDING_LINE_OWNER
    with _PROGRESS_LOCK:
        # Clear incomplete line flag before printing
        spinner.clear_incomplete_line()
        spinner.pause_for_output()
        try:
            current_thread = threading.get_ident()
            # Only print inline if we own the pending line
            if _PENDING_LINE_OWNER == current_thread:
                suffix = f" {extra}" if extra else ""
                print(f" [  \033[32mOK\033[0m  ]{suffix}", flush=True)
                _PENDING_LINE_OWNER = None
            else:
                # Another thread's line or no pending line - print on new line
                if _PENDING_LINE_OWNER is not None:
                    print(flush=True)  # Finish the other thread's line
                suffix = f" {extra}" if extra else ""
                print(
                    f"[{_get_elapsed_secs():12.6f}]  [  \033[32mOK\033[0m  ]{suffix}",
                    flush=True,
                )
                _PENDING_LINE_OWNER = None
        finally:
            spinner.resume_after_output()


def _progress_fail(reason: str = "") -> None:
    """Print FAIL status in Linux init style [FAILED]."""
    global _PENDING_LINE_OWNER
    with _PROGRESS_LOCK:
        # Clear incomplete line flag before printing
        spinner.clear_incomplete_line()
        spinner.pause_for_output()
        try:
            current_thread = threading.get_ident()
            if _PENDING_LINE_OWNER == current_thread:
                suffix = f" ({reason})" if reason else ""
                print(f" [\033[31mFAILED\033[0m]{suffix}", flush=True)
                _PENDING_LINE_OWNER = None
            else:
                if _PENDING_LINE_OWNER is not None:
                    print(flush=True)
                suffix = f" ({reason})" if reason else ""
                print(
                    f"[{_get_elapsed_secs():12.6f}]  [\033[31mFAILED\033[0m]{suffix}",
                    flush=True,
                )
                _PENDING_LINE_OWNER = None
        finally:
            spinner.resume_after_output()


def _progress_skip(reason: str = "") -> None:
    """Print SKIP status in Linux init style [ SKIP ]."""
    global _PENDING_LINE_OWNER
    with _PROGRESS_LOCK:
        # Clear incomplete line flag before printing
        spinner.clear_incomplete_line()
        spinner.pause_for_output()
        try:
            current_thread = threading.get_ident()
            if _PENDING_LINE_OWNER == current_thread:
                suffix = f" ({reason})" if reason else ""
                print(f" [ \033[33mSKIP\033[0m ]{suffix}", flush=True)
                _PENDING_LINE_OWNER = None
            else:
                if _PENDING_LINE_OWNER is not None:
                    print(flush=True)
                suffix = f" ({reason})" if reason else ""
                print(
                    f"[{_get_elapsed_secs():12.6f}]  [ \033[33mSKIP\033[0m ]{suffix}",
                    flush=True,
                )
                _PENDING_LINE_OWNER = None
        finally:
            spinner.resume_after_output()


def enable_vt_mode_if_possible() -> None:
    """!
    @brief Attempt to enable ANSI/VT processing on Windows consoles.
    @details Per the specification, the application should try to enable virtual
    terminal support so both the plain CLI and future TUI renderer can emit
    colorized output. Failures are silently ignored because the feature is
    optional.
    """

    if os.name != "nt":  # pragma: no cover - Windows behaviour only
        return

    try:
        from ctypes import wintypes

        kernel32 = ctypes.windll.kernel32
    except Exception:  # pragma: no cover - import/attribute errors on non-Windows
        return

    for std_handle in (-11, -12):  # STD_OUTPUT_HANDLE, STD_ERROR_HANDLE
        handle = kernel32.GetStdHandle(std_handle)
        if not handle:
            continue
        mode = wintypes.DWORD()
        if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
            kernel32.SetConsoleMode(handle, mode.value | 0x0004)


def ensure_admin_and_relaunch_if_needed() -> None:
    """!
    @brief Request elevation if the current process lacks administrative rights.
    """

    if elevation.is_admin():
        return

    argv = list(sys.argv)
    try:
        executable_path = pathlib.Path(sys.executable).resolve()
        argv0_path = pathlib.Path(argv[0]).resolve() if argv else None
    except Exception:
        executable_path = None
        argv0_path = None

    if executable_path is not None and argv0_path == executable_path:
        argv = argv[1:]

    if not elevation.relaunch_as_admin(argv):
        raise SystemExit("Failed to request elevation via ShellExecuteW.")
    sys.exit(0)


def build_arg_parser() -> argparse.ArgumentParser:
    """!
    @brief Create the top-level argument parser with the specification's surface area.
    @details The parser wires in the mutually exclusive modes and shared options
    defined in the specification so later feature work can hook into the parsed
    values without changing the public CLI signature.
    """

    epilog_text = """\
================================================================================

ODT INSTALLATION PRESETS (use with --odt-install --odt-preset NAME):

  Microsoft 365 (Subscription):
    365-proplus-x64              Microsoft 365 Apps for enterprise (64-bit)
    365-proplus-x86              Microsoft 365 Apps for enterprise (32-bit)
    365-business-x64             Microsoft 365 Apps for business (64-bit)
    365-proplus-visio-project    M365 Apps + Visio + Project (64-bit)
    365-shared-computer          M365 Apps with Shared Computer Licensing

  Office LTSC 2024 (Perpetual):
    office2024-x64               Office LTSC 2024 Professional Plus (64-bit)
    office2024-x86               Office LTSC 2024 Professional Plus (32-bit)
    office2024-standard-x64      Office LTSC 2024 Standard (64-bit)
    ltsc2024-full-x64            Office 2024 + Visio + Project (64-bit)
    ltsc2024-full-x86            Office 2024 + Visio + Project (32-bit)

  Office LTSC 2021 (Perpetual):
    office2021-x64               Office LTSC 2021 Professional Plus (64-bit)
    office2021-x86               Office LTSC 2021 Professional Plus (32-bit)
    office2021-standard-x64      Office LTSC 2021 Standard (64-bit)
    ltsc2021-full-x64            Office 2021 + Visio + Project (64-bit)

  Office 2019 (Perpetual):
    office2019-x64               Office 2019 Professional Plus (64-bit)
    office2019-x86               Office 2019 Professional Plus (32-bit)

  Standalone Products:
    visio-pro-x64                Visio Professional 2024 (64-bit)
    project-pro-x64              Project Professional 2024 (64-bit)

  Clean Presets (No OneDrive/Skype):
    ltsc2024-full-x64-clean      Office 2024 + Visio + Project - No OneDrive/Skype
    ltsc2024-full-x86-clean      Office 2024 + Visio + Project (32-bit) - No bloat
    ltsc2024-x64-clean           Office 2024 ProPlus only - No OneDrive/Skype
    365-proplus-x64-clean        M365 Apps - No OneDrive/Skype
    365-proplus-visio-project-clean  M365 + Visio + Project - No OneDrive/Skype

================================================================================

QUICK INSTALL ALIASES (author shortcuts):

  --goobler    LTSC 2024 + Visio + Project (clean) with pt-pt + en-us
  --pupa       LTSC 2024 ProPlus only (clean) with pt-pt + en-us

================================================================================

QUICK EXAMPLES:

  # Install Office LTSC 2024 + Visio + Project with multiple languages
  office-janitor --odt-install --odt-preset ltsc2024-full-x64 ^
    --odt-language en-us --odt-language es-mx --odt-language pt-br

  # Install clean Office 2024 (no OneDrive/Skype)
  office-janitor --odt-install --odt-preset ltsc2024-full-x64-clean

  # Generate XML config without installing
  office-janitor --odt-build --odt-preset office2024-x64 --odt-output install.xml

  # Remove all Office installations
  office-janitor --auto-all --dry-run  # Preview first!
  office-janitor --auto-all --backup C:\\Backups

  # Diagnose current Office state
  office-janitor --diagnose --plan report.json

================================================================================

SUPPORTED LANGUAGES (common): en-us, de-de, fr-fr, es-es, es-mx, pt-br, pt-pt,
  it-it, ja-jp, ko-kr, zh-cn, zh-tw, ru-ru, pl-pl, nl-nl, ar-sa, he-il, tr-tr

Use --odt-list-languages for the complete list of 60+ supported language codes.
Use --odt-list-products for all available Office product IDs.
Use --odt-list-channels for update channel options.

For legacy OffScrub compatibility, see --offscrub-* options.
Full documentation: https://github.com/yourorg/office-janitor
"""

    parser = argparse.ArgumentParser(
        prog="office-janitor",
        add_help=True,
        description="Detect, uninstall, and scrub Microsoft Office installations.",
        epilog=epilog_text,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    metadata = version.build_info()
    parser.add_argument(
        "-V",
        "--version",
        action="version",
        version=f"{metadata['version']} ({metadata['build']})",
    )

    # -------------------------------------------------------------------------
    # Mode Selection (mutually exclusive)
    # -------------------------------------------------------------------------
    modes = parser.add_mutually_exclusive_group()
    modes.add_argument("--auto-all", action="store_true", help="Run full detection and scrub.")
    modes.add_argument(
        "--target",
        metavar="VER",
        help="Target a specific Office version (2003-2024/365).",
    )
    modes.add_argument(
        "--diagnose",
        action="store_true",
        help="Emit inventory and plan without changes.",
    )
    modes.add_argument(
        "--cleanup-only",
        action="store_true",
        help="Skip uninstalls; clean residue and licensing.",
    )
    modes.add_argument(
        "--repair",
        choices=["quick", "full"],
        metavar="TYPE",
        help="Repair Office C2R (quick|full). Quick runs locally, full uses CDN.",
    )
    modes.add_argument(
        "--repair-config",
        metavar="XML",
        help="Repair/reconfigure using a custom XML configuration file.",
    )

    # -------------------------------------------------------------------------
    # Core Options
    # -------------------------------------------------------------------------
    core_opts = parser.add_argument_group("Core Options")
    core_opts.add_argument(
        "--include",
        metavar="COMPONENTS",
        help="Additional suites/apps to include (visio,project,onenote).",
    )
    core_opts.add_argument(
        "--force", "-f", action="store_true", help="Relax certain guardrails when safe."
    )
    core_opts.add_argument(
        "--allow-unsupported-windows",
        action="store_true",
        help="Permit execution on Windows releases below the supported minimum.",
    )
    core_opts.add_argument(
        "--dry-run",
        "-n",
        action="store_true",
        help="Simulate actions without modifying the system.",
    )
    core_opts.add_argument(
        "--yes",
        "-y",
        action="store_true",
        help="Skip confirmation prompts (assume yes).",
    )
    core_opts.add_argument(
        "--config",
        "-c",
        metavar="JSON",
        help="Load options from a JSON configuration file.",
    )
    core_opts.add_argument(
        "--passes",
        type=int,
        default=None,
        metavar="N",
        help="Number of uninstall passes (default: 1, use higher for stubborn installs).",
    )

    # -------------------------------------------------------------------------
    # Uninstall Method Options
    # -------------------------------------------------------------------------
    uninstall_opts = parser.add_argument_group("Uninstall Method Options")
    uninstall_opts.add_argument(
        "--uninstall-method",
        choices=["auto", "msi", "c2r", "odt", "offscrub"],
        default=None,
        metavar="METHOD",
        help="Preferred uninstall method: auto (detect best), msi (msiexec), c2r (OfficeC2RClient), odt (Office Deployment Tool), offscrub (legacy VBS). Default: auto.",
    )
    uninstall_opts.add_argument(
        "--msi-only",
        action="store_const",
        const="msi",
        dest="uninstall_method",
        help="Only uninstall MSI-based Office products.",
    )
    uninstall_opts.add_argument(
        "--c2r-only",
        action="store_const",
        const="c2r",
        dest="uninstall_method",
        help="Only uninstall Click-to-Run Office products.",
    )
    uninstall_opts.add_argument(
        "--use-odt",
        action="store_const",
        const="odt",
        dest="uninstall_method",
        help="Use Office Deployment Tool (setup.exe) for uninstall.",
    )
    uninstall_opts.add_argument(
        "--force-app-shutdown",
        action="store_true",
        help="Force close running Office applications before uninstall.",
    )
    uninstall_opts.add_argument(
        "--no-force-app-shutdown",
        action="store_true",
        help="Prompt user to close apps instead of forcing shutdown.",
    )
    uninstall_opts.add_argument(
        "--product-code",
        metavar="GUID",
        action="append",
        dest="product_codes",
        help="Specific MSI product code(s) to uninstall. Can be specified multiple times.",
    )
    uninstall_opts.add_argument(
        "--release-id",
        metavar="ID",
        action="append",
        dest="release_ids",
        help="Specific C2R release ID(s) to uninstall (e.g., O365ProPlusRetail). Can be specified multiple times.",
    )

    # -------------------------------------------------------------------------
    # Scrubbing Options
    # -------------------------------------------------------------------------
    scrub_opts = parser.add_argument_group("Scrubbing Options")
    scrub_opts.add_argument(
        "--scrub-level",
        choices=["minimal", "standard", "aggressive", "nuclear"],
        default=None,
        metavar="LEVEL",
        help="Scrub intensity: minimal (uninstall only), standard (+ residue), aggressive (+ deep registry), nuclear (everything). Default: standard.",
    )
    scrub_opts.add_argument(
        "--max-passes",
        type=int,
        default=None,
        metavar="N",
        help="Maximum uninstall/re-detect passes (default: 1). Alias for --passes.",
    )
    scrub_opts.add_argument(
        "--skip-uninstall",
        action="store_true",
        help="Skip uninstall passes entirely; only run cleanup/registry scrubbing.",
    )
    scrub_opts.add_argument(
        "--skip-processes",
        action="store_true",
        help="Skip terminating Office processes before uninstall.",
    )
    scrub_opts.add_argument(
        "--skip-services",
        action="store_true",
        help="Skip stopping Office services before uninstall.",
    )
    scrub_opts.add_argument(
        "--skip-tasks",
        action="store_true",
        help="Skip removing scheduled tasks.",
    )
    scrub_opts.add_argument(
        "--skip-registry",
        action="store_true",
        help="Skip registry cleanup after uninstall.",
    )
    scrub_opts.add_argument(
        "--skip-filesystem",
        action="store_true",
        help="Skip filesystem cleanup after uninstall.",
    )
    scrub_opts.add_argument(
        "--registry-only",
        action="store_true",
        help="Only perform registry cleanup (skip uninstall, filesystem, licensing).",
    )
    scrub_opts.add_argument(
        "--clean-msocache",
        action="store_true",
        help="Also remove MSOCache installation files.",
    )
    scrub_opts.add_argument(
        "--clean-appx",
        action="store_true",
        help="Also remove Office AppX/MSIX packages.",
    )
    scrub_opts.add_argument(
        "--clean-wi-metadata",
        action="store_true",
        help="Clean orphaned Windows Installer metadata.",
    )

    # -------------------------------------------------------------------------
    # License & Activation Options
    # -------------------------------------------------------------------------
    license_opts = parser.add_argument_group("License & Activation Options")
    restore_point_group = license_opts.add_mutually_exclusive_group()
    restore_point_group.add_argument(
        "--restore-point",
        "--create-restore-point",
        action="store_true",
        dest="create_restore_point",
        help="Create a system restore point before scrubbing (default: enabled).",
    )
    restore_point_group.add_argument(
        "--no-restore-point",
        action="store_true",
        help="Skip creating a system restore point.",
    )
    license_opts.add_argument(
        "--no-license", action="store_true", help="Skip license cleanup steps."
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
        help="Aggressively clean all license artifacts (SPP+OSPP+vNext).",
    )

    # -------------------------------------------------------------------------
    # User Data Options
    # -------------------------------------------------------------------------
    data_opts = parser.add_argument_group("User Data Options")
    data_opts.add_argument(
        "--keep-templates",
        action="store_true",
        help="Preserve user templates like normal.dotm.",
    )
    data_opts.add_argument(
        "--keep-user-settings",
        action="store_true",
        help="Preserve user Office settings and customizations.",
    )
    data_opts.add_argument(
        "--delete-user-settings",
        action="store_true",
        help="Remove user Office settings and customizations.",
    )
    data_opts.add_argument(
        "--keep-outlook-data",
        action="store_true",
        help="Preserve Outlook OST/PST files and profiles.",
    )
    data_opts.add_argument(
        "--clean-shortcuts",
        action="store_true",
        help="Remove Office shortcuts from Start Menu and Desktop.",
    )
    data_opts.add_argument(
        "--skip-shortcut-detection",
        action="store_true",
        help="Skip detecting and cleaning orphaned shortcuts.",
    )

    # -------------------------------------------------------------------------
    # Registry Cleanup Options
    # -------------------------------------------------------------------------
    reg_opts = parser.add_argument_group("Registry Cleanup Options")
    reg_opts.add_argument(
        "--clean-addin-registry",
        action="store_true",
        help="Clean Office add-in registry entries.",
    )
    reg_opts.add_argument(
        "--clean-com-registry",
        action="store_true",
        help="Clean orphaned COM/ActiveX registrations.",
    )
    reg_opts.add_argument(
        "--clean-shell-extensions",
        action="store_true",
        help="Clean orphaned shell extensions.",
    )
    reg_opts.add_argument(
        "--clean-typelibs",
        action="store_true",
        help="Clean orphaned type libraries.",
    )
    reg_opts.add_argument(
        "--clean-protocol-handlers",
        action="store_true",
        help="Clean Office protocol handlers (ms-word:, ms-excel:, etc.).",
    )
    reg_opts.add_argument(
        "--remove-vba",
        action="store_true",
        help="Remove VBA-only package and related registry entries.",
    )

    # -------------------------------------------------------------------------
    # Output & Logging Options
    # -------------------------------------------------------------------------
    output_opts = parser.add_argument_group("Output & Logging Options")
    output_opts.add_argument(
        "--plan", metavar="OUT", help="Write the computed action plan to a JSON file."
    )
    output_opts.add_argument(
        "--logdir", metavar="DIR", help="Directory for human/JSONL log output."
    )
    output_opts.add_argument(
        "--backup", metavar="DIR", help="Destination for registry/file backups."
    )
    output_opts.add_argument(
        "--timeout", metavar="SEC", type=int, help="Per-step timeout in seconds."
    )
    output_opts.add_argument(
        "--quiet", "-q", action="store_true", help="Minimal console output (errors only)."
    )
    output_opts.add_argument(
        "--json", action="store_true", help="Mirror structured events to stdout."
    )
    output_opts.add_argument(
        "--verbose",
        "-v",
        action="count",
        default=0,
        help="Increase output verbosity (-v, -vv, -vvv).",
    )

    # -------------------------------------------------------------------------
    # TUI Options
    # -------------------------------------------------------------------------
    tui_opts = parser.add_argument_group("TUI Options")
    tui_opts.add_argument("--tui", action="store_true", help="Force the interactive text UI mode.")
    tui_opts.add_argument("--no-color", action="store_true", help="Disable ANSI color codes.")
    tui_opts.add_argument(
        "--tui-compact",
        action="store_true",
        help="Use a compact TUI layout for small consoles.",
    )
    tui_opts.add_argument(
        "--tui-refresh",
        metavar="MS",
        type=int,
        help="Refresh interval for the TUI renderer in milliseconds.",
    )
    tui_opts.add_argument(
        "--limited-user",
        action="store_true",
        help="Run detection and uninstall stages under a limited user token when possible.",
    )

    # -------------------------------------------------------------------------
    # Retry & Resilience Options
    # -------------------------------------------------------------------------
    retry_opts = parser.add_argument_group("Retry & Resilience Options")
    retry_opts.add_argument(
        "--retries",
        type=int,
        default=None,
        metavar="N",
        help="Number of retry attempts per step (default: 4).",
    )
    retry_opts.add_argument(
        "--retry-delay",
        type=int,
        default=3,
        metavar="SEC",
        help="Base delay between retries in seconds (default: 3).",
    )
    retry_opts.add_argument(
        "--retry-delay-max",
        type=int,
        default=30,
        metavar="SEC",
        help="Maximum delay between retries in seconds (default: 30).",
    )
    retry_opts.add_argument(
        "--no-reboot",
        action="store_true",
        help="Suppress reboot recommendations even if services require it.",
    )
    retry_opts.add_argument(
        "--offline",
        action="store_true",
        help="Run in offline mode (no network access for downloads).",
    )

    # -------------------------------------------------------------------------
    # ODT Build Options (XML Configuration Generator)
    # -------------------------------------------------------------------------
    odt_build_opts = parser.add_argument_group(
        "ODT Build Options",
        "Generate Office Deployment Tool XML configurations for installation.",
    )
    odt_build_opts.add_argument(
        "--odt-build",
        action="store_true",
        help="Generate an ODT XML configuration file instead of running scrub operations.",
    )
    odt_build_opts.add_argument(
        "--odt-install",
        action="store_true",
        help="Install Office using ODT with the specified configuration (requires elevation).",
    )
    odt_build_opts.add_argument(
        "--odt-preset",
        metavar="NAME",
        help="Use a predefined ODT installation preset (use --odt-list-presets to see options).",
    )
    odt_build_opts.add_argument(
        "--odt-product",
        metavar="ID",
        action="append",
        dest="odt_products",
        help="Product ID to include in ODT config (can be repeated). Use --odt-list-products to see options.",
    )
    odt_build_opts.add_argument(
        "--odt-language",
        metavar="CODE",
        action="append",
        dest="odt_languages",
        help="Language code for ODT config (can be repeated, default: en-us).",
    )
    odt_build_opts.add_argument(
        "--odt-arch",
        choices=["32", "64"],
        default="64",
        metavar="BITS",
        help="Architecture for ODT config: 32 or 64 (default: 64).",
    )
    odt_build_opts.add_argument(
        "--odt-channel",
        metavar="CHANNEL",
        help="Update channel for ODT config (e.g., Current, PerpetualVL2024).",
    )
    odt_build_opts.add_argument(
        "--odt-output",
        metavar="FILE",
        help="Output path for generated ODT XML (default: stdout).",
    )
    odt_build_opts.add_argument(
        "--odt-shared-computer",
        action="store_true",
        help="Enable shared computer licensing in ODT config.",
    )
    odt_build_opts.add_argument(
        "--odt-remove-msi",
        action="store_true",
        help="Include RemoveMSI element to remove existing MSI Office.",
    )
    odt_build_opts.add_argument(
        "--odt-exclude-app",
        metavar="APP",
        action="append",
        dest="odt_exclude_apps",
        help="App to exclude from installation (e.g., OneDrive, Teams). Can be repeated.",
    )
    odt_build_opts.add_argument(
        "--odt-include-visio",
        action="store_true",
        help="Include Visio Professional in the ODT configuration.",
    )
    odt_build_opts.add_argument(
        "--odt-include-project",
        action="store_true",
        help="Include Project Professional in the ODT configuration.",
    )
    odt_build_opts.add_argument(
        "--odt-removal",
        action="store_true",
        help="Generate a removal XML instead of installation XML.",
    )
    odt_build_opts.add_argument(
        "--odt-download",
        metavar="PATH",
        help="Generate a download XML with the specified local path for offline files.",
    )
    odt_build_opts.add_argument(
        "--odt-list-products",
        action="store_true",
        help="List all available ODT product IDs and exit.",
    )
    odt_build_opts.add_argument(
        "--odt-list-presets",
        action="store_true",
        help="List all available ODT installation presets and exit.",
    )
    odt_build_opts.add_argument(
        "--odt-list-channels",
        action="store_true",
        help="List all available ODT update channels and exit.",
    )
    odt_build_opts.add_argument(
        "--odt-list-languages",
        action="store_true",
        help="List all supported language codes and exit.",
    )

    # -------------------------------------------------------------------------
    # Author Quick Install Aliases
    # -------------------------------------------------------------------------
    author_opts = parser.add_argument_group(
        "Quick Install Aliases",
        "Author-defined shortcuts for common Office installations.",
    )
    author_opts.add_argument(
        "--goobler",
        action="store_true",
        help="Install LTSC 2024 + Visio + Project (clean, no OneDrive/Skype) with pt-pt and en-us.",
    )
    author_opts.add_argument(
        "--pupa",
        action="store_true",
        help="Install LTSC 2024 ProPlus only (clean, no OneDrive/Skype) with pt-pt and en-us.",
    )

    # -------------------------------------------------------------------------
    # Repair Options
    # -------------------------------------------------------------------------
    repair_opts = parser.add_argument_group("Repair Options")
    repair_opts.add_argument(
        "--repair-culture",
        metavar="LANG",
        default="en-us",
        help="Language/culture code for repair (default: en-us).",
    )
    repair_opts.add_argument(
        "--repair-platform",
        choices=["x86", "x64"],
        metavar="ARCH",
        help="Architecture for repair (auto-detected if not specified).",
    )
    repair_opts.add_argument(
        "--repair-visible",
        action="store_true",
        help="Show repair UI instead of running silently.",
    )
    repair_opts.add_argument(
        "--repair-timeout",
        type=int,
        default=3600,
        metavar="SEC",
        help="Timeout for repair operations in seconds (default: 3600).",
    )

    # -------------------------------------------------------------------------
    # OEM Configuration Presets
    # -------------------------------------------------------------------------
    oem_opts = parser.add_argument_group("OEM Configuration Presets")
    oem_configs = oem_opts.add_mutually_exclusive_group()
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
        help="Remove all Office C2R products (alias for --oem-config full-removal).",
    )
    oem_configs.add_argument(
        "--c2r-repair-quick",
        action="store_const",
        const="quick-repair",
        dest="oem_config",
        help="Quick repair Office C2R (alias for --oem-config quick-repair).",
    )
    oem_configs.add_argument(
        "--c2r-repair-full",
        action="store_const",
        const="full-repair",
        dest="oem_config",
        help="Full online repair Office C2R (alias for --oem-config full-repair).",
    )
    oem_configs.add_argument(
        "--c2r-proplus",
        action="store_const",
        const="proplus-x64",
        dest="oem_config",
        help="Repair Office 365 ProPlus x64 (alias for --oem-config proplus-x64).",
    )
    oem_configs.add_argument(
        "--c2r-business",
        action="store_const",
        const="business-x64",
        dest="oem_config",
        help="Repair Microsoft 365 Business x64 (alias for --oem-config business-x64).",
    )

    # -------------------------------------------------------------------------
    # OffScrub Legacy Compatibility Flags
    # -------------------------------------------------------------------------
    offscrub_opts = parser.add_argument_group(
        "OffScrub Legacy Compatibility", "Flags for compatibility with legacy OffScrub VBS scripts."
    )
    offscrub_opts.add_argument(
        "--offscrub-all",
        action="store_true",
        help="OffScrub /ALL: Remove all detected Office products.",
    )
    offscrub_opts.add_argument(
        "--offscrub-ose",
        action="store_true",
        help="OffScrub /OSE: Fix OSE service configuration before uninstall.",
    )
    offscrub_opts.add_argument(
        "--offscrub-offline",
        action="store_true",
        help="OffScrub /OFFLINE: Mark C2R config as offline mode.",
    )
    offscrub_opts.add_argument(
        "--offscrub-quiet",
        action="store_true",
        help="OffScrub /QUIET: Reduce human log verbosity.",
    )
    offscrub_opts.add_argument(
        "--offscrub-test-rerun",
        action="store_true",
        help="OffScrub /TR: Run uninstall passes twice (test rerun).",
    )
    offscrub_opts.add_argument(
        "--offscrub-bypass",
        action="store_true",
        help="OffScrub /BYPASS: Bypass certain safety checks.",
    )
    offscrub_opts.add_argument(
        "--offscrub-fast-remove",
        action="store_true",
        help="OffScrub /FASTREMOVE: Skip verification probes after uninstall.",
    )
    offscrub_opts.add_argument(
        "--offscrub-scan-components",
        action="store_true",
        help="OffScrub /SCANCOMPONENTS: Scan Windows Installer components.",
    )
    offscrub_opts.add_argument(
        "--offscrub-return-error",
        action="store_true",
        help="OffScrub /RETERRORSUCCESS: Return error codes instead of success on partial.",
    )

    # -------------------------------------------------------------------------
    # Advanced Options
    # -------------------------------------------------------------------------
    adv_opts = parser.add_argument_group("Advanced Options")
    adv_opts.add_argument(
        "--skip-preflight",
        action="store_true",
        help="Skip preflight safety checks (use with caution).",
    )
    adv_opts.add_argument(
        "--skip-backup",
        action="store_true",
        help="Skip creating registry and file backups.",
    )
    adv_opts.add_argument(
        "--skip-verification",
        action="store_true",
        help="Skip verification probes after uninstall.",
    )
    adv_opts.add_argument(
        "--schedule-reboot",
        action="store_true",
        help="Schedule a reboot after completion if recommended.",
    )
    adv_opts.add_argument(
        "--no-schedule-delete",
        action="store_true",
        help="Don't use MoveFileEx for locked file deletion on reboot.",
    )
    adv_opts.add_argument(
        "--msiexec-args",
        metavar="ARGS",
        help="Additional arguments to pass to msiexec (e.g., '/l*v log.txt').",
    )
    adv_opts.add_argument(
        "--c2r-args",
        metavar="ARGS",
        help="Additional arguments to pass to OfficeC2RClient.exe.",
    )
    adv_opts.add_argument(
        "--odt-args",
        metavar="ARGS",
        help="Additional arguments to pass to ODT setup.exe.",
    )

    return parser


def _resolve_log_directory(candidate: str | None) -> pathlib.Path:
    """!
    @brief Determine the log directory path using specification defaults when unspecified.
    """

    if candidate:
        return pathlib.Path(candidate).expanduser().resolve()
    default_dir = fs_tools.get_default_log_directory()
    expanded = default_dir.expanduser()
    try:
        return expanded.resolve()
    except Exception:
        return expanded


def _bootstrap_logging(
    args: argparse.Namespace,
) -> tuple[logging.Logger, logging.Logger]:
    """!
    @brief Initialize human and machine loggers using :mod:`logging_ext` helpers.
    @returns A tuple of configured human and machine loggers.
    """

    logdir = _resolve_log_directory(getattr(args, "logdir", None))
    logdir.mkdir(parents=True, exist_ok=True)
    args.logdir = str(logdir)
    human_logger, machine_logger = logging_ext.setup_logging(
        logdir,
        json_to_stdout=getattr(args, "json", False),
    )
    args.human_logger = human_logger
    args.machine_logger = machine_logger
    if getattr(args, "quiet", False):
        human_logger.setLevel(logging.ERROR)
    return human_logger, machine_logger


def _handle_shutdown_signal(signum: int, frame: object) -> None:
    """!
    @brief Handle Ctrl+C (SIGINT) for clean shutdown.
    """
    sig_name = signal.Signals(signum).name if hasattr(signal, "Signals") else str(signum)
    elapsed = _get_elapsed_secs()
    print(flush=True)  # Newline after any partial output
    print(f"[{elapsed:12.6f}] " + "=" * 50, flush=True)
    print(
        f"[{elapsed:12.6f}] \033[33mShutting down...\033[0m (received {sig_name})",
        flush=True,
    )
    print(f"[{elapsed:12.6f}] Cleanup in progress, please wait...", flush=True)
    print(f"[{elapsed:12.6f}] " + "=" * 50, flush=True)
    sys.exit(130)  # Standard exit code for SIGINT


def _print_shutdown_message() -> None:
    """Print a clean shutdown message with timestamp."""
    elapsed = _get_elapsed_secs()
    print(flush=True)  # Newline after any partial output
    print(f"[{elapsed:12.6f}] " + "=" * 50, flush=True)
    print(
        f"[{elapsed:12.6f}] \033[33mShutting down...\033[0m (keyboard interrupt)",
        flush=True,
    )
    print(f"[{elapsed:12.6f}] Cleanup in progress, please wait...", flush=True)
    print(f"[{elapsed:12.6f}] " + "=" * 50, flush=True)


def _load_config_file(config_path: str | None) -> dict[str, object]:
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
            print(f"Error: Configuration file must contain a JSON object: {path}", file=sys.stderr)
            raise SystemExit(1)
        return config
    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON in configuration file: {path}\n{e}", file=sys.stderr)
        raise SystemExit(1)
    except OSError as e:
        print(f"Error: Cannot read configuration file: {path}\n{e}", file=sys.stderr)
        raise SystemExit(1)


def main(argv: Iterable[str] | None = None, *, start_time: float | None = None) -> int:
    """!
    @brief Entry point invoked by the shim and PyInstaller bundle.
    @param start_time Optional startup timestamp from entry point for continuous timing.
    @returns Process exit code integer.
    """
    global _MAIN_START_TIME
    # Use provided start_time for continuous timestamps, or start fresh
    _MAIN_START_TIME = start_time if start_time is not None else time.perf_counter()

    # Check for elevation marker and set environment flag
    argv_list = list(argv) if argv is not None else None
    if argv_list is None:
        argv_list = sys.argv[1:]
    if argv_list and argv_list[0] == "--_elevated-marker":
        os.environ["OFFICE_JANITOR_ELEVATED"] = "1"
        argv_list = argv_list[1:]  # Remove the marker from args

    # Install signal handler for clean Ctrl+C shutdown
    signal.signal(signal.SIGINT, _handle_shutdown_signal)
    if hasattr(signal, "SIGBREAK"):  # Windows-specific
        signal.signal(signal.SIGBREAK, _handle_shutdown_signal)

    try:
        exit_code = _main_impl(argv_list)
    except KeyboardInterrupt:
        _print_shutdown_message()
        exit_code = 130  # Standard exit code for SIGINT

    # Keep console open if this was an auto-elevated process
    elevation.pause_if_elevated(exit_code)
    return exit_code


def _main_impl(argv: Iterable[str] | None = None) -> int:
    """!
    @brief Implementation of main() wrapped with KeyboardInterrupt handling.
    """
    _progress("=" * 60)
    _progress("Office Janitor - Main Entry Point")
    _progress("=" * 60)

    # Phase 0: Pre-parse for ODT listing commands (no elevation needed)
    parser = build_arg_parser()
    args = parser.parse_args(list(argv) if argv is not None else None)

    # Handle ODT listing commands early (before elevation check)
    if _handle_odt_list_commands(args):
        return 0

    # Phase 1: Elevation check
    _progress("Phase 1: Checking administrative privileges...", newline=False)
    try:
        ensure_admin_and_relaunch_if_needed()
        _progress_ok("elevated")
    except SystemExit:
        _progress_fail("elevation required")
        raise

    # Phase 2: Console setup
    _progress("Phase 2: Configuring console...", newline=False)
    enable_vt_mode_if_possible()
    _progress_ok("VT mode enabled")

    # Handle ODT build commands (after elevation since they write files)
    if _handle_odt_build_commands(args):
        return 0

    # Phase 3: Timeout configuration
    _progress("Phase 3: Configuring execution timeout...", newline=False)
    timeout_val = getattr(args, "timeout", None)
    exec_utils.set_global_timeout(timeout_val)
    _progress_ok(f"{timeout_val}s" if timeout_val else "default")

    # Phase 5: Logging bootstrap
    _progress("Phase 5: Initializing logging subsystem...", newline=False)
    human_log, machine_log = _bootstrap_logging(args)
    if getattr(args, "quiet", False):
        human_log.setLevel(logging.ERROR)
    _progress_ok(f"logdir={getattr(args, 'logdir', 'default')}")

    # Phase 6: Mode determination
    _progress("Phase 6: Determining operation mode...", newline=False)
    mode = _determine_mode(args)
    _progress_ok(mode)

    # Log startup event
    _progress("Emitting startup telemetry...")
    machine_log.info(
        "startup",
        extra={
            "event": "startup",
            "data": {"mode": mode, "dry_run": bool(getattr(args, "dry_run", False))},
        },
    )

    # Phase 7: App state construction
    _progress("Phase 7: Building application state...", newline=False)
    app_state = _build_app_state(args, human_log, machine_log)
    _progress_ok()

    # Repair mode handling - separate from standard detection/scrub flow
    if mode.startswith("repair:"):
        _progress("Entering repair mode...")
        return _handle_repair_mode(args, mode, human_log, machine_log)

    # OEM config mode handling - execute bundled XML configurations
    if mode.startswith("oem-config:"):
        _progress("Entering OEM configuration mode...")
        return _handle_oem_config_mode(args, mode, human_log, machine_log)

    # Interactive mode handling
    if mode == "interactive":
        _progress("Entering interactive mode...")
        if getattr(args, "tui", False):
            _progress("Launching TUI (forced via --tui)...")
            tui.run_tui(app_state)
        else:
            tui_candidate = _should_use_tui(args)
            if tui_candidate:
                _progress("Launching TUI (auto-detected)...")
                tui.run_tui(app_state)
            else:
                _progress("Launching CLI interface...")
                ui.run_cli(app_state)
        _progress("Interactive session complete.")
        return 0

    _progress("-" * 60)
    _progress(f"Running in non-interactive mode: {mode}")
    _progress("-" * 60)

    # Start spinner for non-interactive mode
    spinner.start_spinner_thread()

    # Phase 8: Detection
    spinner.set_task("Detecting Office installations")
    _progress("Phase 8: Running Office detection...")
    logdir_path = pathlib.Path(getattr(args, "logdir", _resolve_log_directory(None))).expanduser()
    limited_flag = bool(getattr(args, "limited_user", False))
    if limited_flag:
        _progress("Using limited user token for detection", indent=1)
    inventory = _run_detection(
        machine_log,
        logdir_path,
        limited_user=limited_flag or None,
    )
    item_count = sum(len(v) if hasattr(v, "__len__") else 0 for v in inventory.values())
    _progress(f"Detection complete: {item_count} items found", indent=1)

    # Phase 9: Plan generation
    spinner.set_task("Building execution plan")
    _progress("Phase 9: Building execution plan...")
    options = _collect_plan_options(args, mode)
    _progress(
        f"Plan options: dry_run={options.get('dry_run')}, force={options.get('force')}",
        indent=1,
    )
    generated_plan = plan_module.build_plan(inventory, options)
    _progress(f"Generated {len(generated_plan)} plan steps", indent=1)

    # Phase 10: Safety checks (warnings only - never block execution)
    spinner.set_task("Performing safety checks")
    _progress("Phase 10: Performing preflight safety checks...")
    try:
        safety.perform_preflight_checks(generated_plan)
        _progress("All preflight checks passed", indent=1, newline=False)
        _progress_ok()
    except ValueError as err:
        _progress(f"Warning: {err}", indent=1, newline=False)
        _progress_skip("non-fatal")
        human_log.warning("Preflight safety warning: %s", err)
    except Exception as err:  # noqa: BLE001
        _progress(f"Warning: {err}", indent=1, newline=False)
        _progress_skip("non-fatal")
        human_log.warning("Unexpected preflight warning: %s", err)

    # Phase 11: Artifacts
    spinner.set_task("Writing plan artifacts")
    _progress("Phase 11: Writing plan artifacts...")
    _handle_plan_artifacts(args, generated_plan, inventory, human_log, mode)

    if mode == "diagnose":
        spinner.clear_task()
        spinner.stop_spinner_thread()
        _progress("=" * 60)
        _progress("Diagnostics complete - no actions executed")
        _progress("=" * 60)
        human_log.info("Diagnostics complete; plan written and no actions executed.")
        return 0

    # Phase 12: User confirmation
    spinner.set_task("Awaiting confirmation")
    _progress("Phase 12: Requesting user confirmation...")
    scrub_dry_run = bool(getattr(args, "dry_run", False))
    is_auto_all = mode == "auto-all"
    proceed = confirm.request_scrub_confirmation(
        dry_run=scrub_dry_run,
        force=bool(getattr(args, "force", False)) or is_auto_all,
    )
    if not proceed:
        spinner.clear_task()
        spinner.stop_spinner_thread()
        _progress("User declined confirmation - aborting")
        human_log.info("Scrub cancelled by user confirmation prompt.")
        machine_log.info(
            "scrub.cancelled",
            extra={
                "event": "scrub.cancelled",
                "data": {"reason": "user_declined", "dry_run": scrub_dry_run},
            },
        )
        return 0
    _progress("Confirmation received", indent=1)

    # Phase 13: Runtime guards (warnings only - never block execution)
    spinner.set_task("Enforcing runtime guards")
    _progress("Phase 13: Enforcing runtime guards...")
    try:
        _enforce_runtime_guards(options, dry_run=scrub_dry_run)
        _progress("Runtime guards passed", indent=1, newline=False)
        _progress_ok()
    except ValueError as err:
        _progress(f"Warning: {err}", indent=1, newline=False)
        _progress_skip("non-fatal")
        human_log.warning("Runtime guard warning: %s", err)
    except Exception as err:  # noqa: BLE001
        _progress(f"Warning: {err}", indent=1, newline=False)
        _progress_skip("non-fatal")
        human_log.warning("Unexpected runtime warning: %s", err)

    # Phase 14: Plan execution
    spinner.set_task("Executing scrub plan")
    _progress("=" * 60)
    _progress(f"Phase 14: Executing plan ({'DRY RUN' if scrub_dry_run else 'LIVE'})...")
    _progress("=" * 60)
    fatal_error: str | None = None
    try:
        scrub.execute_plan(generated_plan, dry_run=scrub_dry_run, start_time=_MAIN_START_TIME)
    except KeyboardInterrupt:
        _progress("Execution interrupted by user", indent=1, newline=False)
        _progress_skip("cancelled")
        human_log.warning("Plan execution cancelled by user")
        raise
    except (OSError, PermissionError) as err:
        # These are potentially fatal - system-level failures
        fatal_error = str(err)
        _progress(f"Fatal error: {err}", indent=1, newline=False)
        _progress_fail()
        human_log.error("Fatal execution error: %s", err)
    except Exception as err:  # noqa: BLE001
        # Log but don't treat as fatal - execution may have partially succeeded
        _progress(f"Error during execution: {err}", indent=1, newline=False)
        _progress_fail()
        human_log.exception("Error during plan execution (non-fatal)")

    # Stop spinner before final messages
    spinner.clear_task()
    spinner.stop_spinner_thread()

    _progress("=" * 60)
    if fatal_error:
        _progress(f"Execution failed after {_get_elapsed_secs():.3f}s")
        _progress("=" * 60)
        return 1
    _progress(f"Execution complete in {_get_elapsed_secs():.3f}s")
    _progress("=" * 60)
    return 0


def _handle_odt_list_commands(args: argparse.Namespace) -> bool:
    """!
    @brief Handle ODT listing commands that don't require elevation.
    @details Processes --odt-list-* commands which are read-only operations
    that can run without administrative privileges.
    @param args Parsed command-line arguments.
    @returns True if a list command was handled (caller should exit), False otherwise.
    """
    if getattr(args, "odt_list_products", False):
        print("\nAvailable Office Products for ODT Configuration:")
        print("=" * 80)
        for product in odt_build.list_products():
            channels = ", ".join(product.get("channels", []))  # type: ignore[arg-type]
            print(f"\n  {product['id']}")
            print(f"      Name: {product['name']}")
            print(f"      Channels: {channels}")
            print(f"      {product['description']}")
        return True

    if getattr(args, "odt_list_presets", False):
        print("\nAvailable Installation Presets:")
        print("=" * 80)
        for preset in odt_build.list_presets():
            products = ", ".join(preset.get("products", []))  # type: ignore[arg-type]
            print(f"\n  {preset['name']}")
            print(f"      Products: {products}")
            print(f"      Architecture: {preset['architecture']}, Channel: {preset['channel']}")
            print(f"      {preset['description']}")
        return True

    if getattr(args, "odt_list_channels", False):
        print("\nAvailable Update Channels:")
        print("=" * 60)
        for ch in odt_build.list_channels():
            print(f"  {ch['name']:<30} {ch['value']}")
        return True

    if getattr(args, "odt_list_languages", False):
        print("\nSupported Language Codes:")
        print("=" * 60)
        langs = odt_build.list_languages()
        for i in range(0, len(langs), 4):
            row = langs[i : i + 4]
            print("  " + "  ".join(f"{lang:<14}" for lang in row))
        return True

    return False


def _handle_odt_build_commands(args: argparse.Namespace) -> bool:
    """!
    @brief Handle ODT build and configuration generation commands.
    @details Processes --odt-build, --odt-install, --odt-removal, --odt-download,
    and author aliases (--goobler, --pupa).
    These commands may write files and run after elevation.
    @param args Parsed command-line arguments.
    @returns True if an ODT command was handled (caller should exit), False otherwise.
    """
    # Handle author quick install aliases
    if getattr(args, "goobler", False):
        return _run_author_install(
            args,
            preset="ltsc2024-full-x64-clean",
            languages=["pt-pt", "en-us"],
            name="Goobler",
        )

    if getattr(args, "pupa", False):
        return _run_author_install(
            args,
            preset="ltsc2024-x64-clean",
            languages=["pt-pt", "en-us"],
            name="Pupa",
        )

    # Handle install command (actually runs ODT)
    if getattr(args, "odt_install", False):
        return _run_odt_install(args)

    # Handle build command (generates XML)
    if getattr(args, "odt_build", False):
        return _build_odt_config(args)

    # Handle removal XML generation
    if getattr(args, "odt_removal", False):
        return _build_odt_removal(args)

    # Handle download XML generation
    if getattr(args, "odt_download", None):
        return _build_odt_download(args)

    return False


def _run_odt_install(args: argparse.Namespace) -> bool:
    """!
    @brief Run ODT to install Office with the specified configuration.
    @param args Parsed command-line arguments.
    @returns True (command was handled).
    """
    dry_run = getattr(args, "dry_run", False)

    try:
        preset = getattr(args, "odt_preset", None)
        products = getattr(args, "odt_products", None)
        languages = getattr(args, "odt_languages", None) or ["en-us"]

        if preset:
            config = odt_build.ODTConfig.from_preset(preset, languages)
        elif products:
            arch = (
                odt_build.Architecture.X64
                if getattr(args, "odt_arch", "64") == "64"
                else odt_build.Architecture.X86
            )

            channel_arg = getattr(args, "odt_channel", None)
            channel = odt_build.UpdateChannel.CURRENT
            if channel_arg:
                for ch in odt_build.UpdateChannel:
                    if ch.name.lower() == channel_arg.lower().replace("-", "_"):
                        channel = ch
                        break
                    if ch.value.lower() == channel_arg.lower():
                        channel = ch
                        break

            exclude_apps = getattr(args, "odt_exclude_apps", None) or []
            product_configs = [
                odt_build.ProductConfig(pid, languages=languages, exclude_apps=exclude_apps)
                for pid in products
            ]

            if getattr(args, "odt_include_visio", False):
                product_configs.append(
                    odt_build.ProductConfig("VisioProRetail", languages=languages)
                )
            if getattr(args, "odt_include_project", False):
                product_configs.append(
                    odt_build.ProductConfig("ProjectProRetail", languages=languages)
                )

            config = odt_build.ODTConfig(
                products=product_configs,
                architecture=arch,
                channel=channel,
                shared_computer_licensing=getattr(args, "odt_shared_computer", False),
                remove_msi=getattr(args, "odt_remove_msi", False),
            )
        else:
            _progress("No preset or products specified for installation.")
            _progress("  Use --odt-preset or --odt-product to specify what to install.", indent=1)
            _progress("  Use --odt-list-presets or --odt-list-products for options.", indent=1)
            return True

        # Show what we're about to install using consistent logging format
        _progress("-" * 60)
        _progress("ODT Installation")
        _progress("-" * 60)
        _progress(f"  Products: {', '.join(p.product_id for p in config.products)}")
        _progress(
            f"  Languages: {', '.join(config.products[0].languages) if config.products else 'none'}"
        )
        _progress(f"  Architecture: {config.architecture.value}-bit")
        _progress(f"  Channel: {config.channel.value}")
        _progress("-" * 60)

        # Run installation - spinner handles progress display, don't use incomplete line
        result = odt_build.run_odt_install(config, dry_run=dry_run)

        if result.success:
            _progress_ok(f"{result.duration:.1f}s")
            _progress("Installation complete")
        else:
            _progress_fail(f"exit {result.return_code}")
            if result.stderr:
                _progress(f"  Error: {result.stderr}", indent=1)
            if result.config_path:
                _progress(f"  Config preserved: {result.config_path}", indent=1)

        return True

    except ValueError as e:
        _progress_fail(str(e))
        return True
    except FileNotFoundError as e:
        _progress_fail(str(e))
        return True


def _run_author_install(
    args: argparse.Namespace,
    *,
    preset: str,
    languages: list[str],
    name: str,
) -> bool:
    """!
    @brief Run a pre-configured author install alias.
    @param args Parsed command-line arguments.
    @param preset The preset name to use.
    @param languages List of language codes.
    @param name Display name for the alias.
    @returns True (command was handled).
    """
    dry_run = getattr(args, "dry_run", False)

    try:
        config = odt_build.ODTConfig.from_preset(preset, languages)

        # Show what we're about to install using consistent logging format
        _progress("-" * 60)
        _progress(f"Quick Install: {name}")
        _progress("-" * 60)
        _progress(f"  Preset: {preset}")
        _progress(f"  Products: {', '.join(p.product_id for p in config.products)}")
        _progress(f"  Languages: {', '.join(languages)}")
        _progress(f"  Architecture: {config.architecture.value}-bit")
        _progress(f"  Channel: {config.channel.value}")
        if config.products and config.products[0].exclude_apps:
            _progress(f"  Excluded: {', '.join(config.products[0].exclude_apps)}")
        _progress("-" * 60)

        # Run installation - spinner handles progress display, don't use incomplete line
        result = odt_build.run_odt_install(config, dry_run=dry_run)

        if result.success:
            _progress_ok(f"{result.duration:.1f}s")
            _progress(f"Installation complete: {name}")
        else:
            _progress_fail(f"exit {result.return_code}")
            if result.stderr:
                _progress(f"  Error: {result.stderr}", indent=1)
            if result.config_path:
                _progress(f"  Config preserved: {result.config_path}", indent=1)

        return True

    except ValueError as e:
        _progress_fail(str(e))
        return True
    except FileNotFoundError as e:
        _progress_fail(str(e))
        return True


def _build_odt_config(args: argparse.Namespace) -> bool:
    """!
    @brief Build ODT installation XML configuration.
    @param args Parsed command-line arguments.
    @returns True (command was handled).
    """
    try:
        preset = getattr(args, "odt_preset", None)
        products = getattr(args, "odt_products", None)
        languages = getattr(args, "odt_languages", None) or ["en-us"]

        if preset:
            # Use preset configuration
            config = odt_build.ODTConfig.from_preset(preset, languages)
        elif products:
            # Build custom configuration from product list
            arch = (
                odt_build.Architecture.X64
                if getattr(args, "odt_arch", "64") == "64"
                else odt_build.Architecture.X86
            )

            # Determine channel
            channel_arg = getattr(args, "odt_channel", None)
            channel = odt_build.UpdateChannel.CURRENT
            if channel_arg:
                # Try to match channel by name or value
                for ch in odt_build.UpdateChannel:
                    if ch.name.lower() == channel_arg.lower().replace("-", "_"):
                        channel = ch
                        break
                    if ch.value.lower() == channel_arg.lower():
                        channel = ch
                        break

            exclude_apps = getattr(args, "odt_exclude_apps", None) or []
            product_configs = [
                odt_build.ProductConfig(pid, languages=languages, exclude_apps=exclude_apps)
                for pid in products
            ]

            # Add optional products
            if getattr(args, "odt_include_visio", False):
                product_configs.append(
                    odt_build.ProductConfig("VisioProRetail", languages=languages)
                )
            if getattr(args, "odt_include_project", False):
                product_configs.append(
                    odt_build.ProductConfig("ProjectProRetail", languages=languages)
                )

            config = odt_build.ODTConfig(
                products=product_configs,
                architecture=arch,
                channel=channel,
                shared_computer_licensing=getattr(args, "odt_shared_computer", False),
                remove_msi=getattr(args, "odt_remove_msi", False),
            )
        else:
            # No preset or products specified - use default M365 ProPlus
            print("No preset or products specified. Use --odt-preset or --odt-product.")
            print("Use --odt-list-presets or --odt-list-products to see available options.")
            return True

        xml_output = odt_build.build_xml(config)

        output_path = getattr(args, "odt_output", None)
        if output_path:
            pathlib.Path(output_path).write_text(xml_output, encoding="utf-8")
            print(f"ODT configuration written to: {output_path}")
        else:
            print(xml_output)

    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)
        return True

    return True


def _build_odt_removal(args: argparse.Namespace) -> bool:
    """!
    @brief Build ODT removal XML configuration.
    @param args Parsed command-line arguments.
    @returns True (command was handled).
    """
    product_ids = getattr(args, "odt_products", None)
    force_shutdown = not getattr(args, "no_force_app_shutdown", False)
    remove_msi = getattr(args, "odt_remove_msi", False)

    xml_output = odt_build.build_removal_xml(
        remove_all=not product_ids,
        product_ids=product_ids,
        force_app_shutdown=force_shutdown,
        remove_msi=remove_msi,
    )

    output_path = getattr(args, "odt_output", None)
    if output_path:
        pathlib.Path(output_path).write_text(xml_output, encoding="utf-8")
        print(f"ODT removal configuration written to: {output_path}")
    else:
        print(xml_output)

    return True


def _build_odt_download(args: argparse.Namespace) -> bool:
    """!
    @brief Build ODT download XML configuration.
    @param args Parsed command-line arguments.
    @returns True (command was handled).
    """
    download_path = getattr(args, "odt_download", None)
    if not download_path:
        return False

    try:
        preset = getattr(args, "odt_preset", None)
        products = getattr(args, "odt_products", None)
        languages = getattr(args, "odt_languages", None) or ["en-us"]

        if preset:
            config = odt_build.ODTConfig.from_preset(preset, languages)
        elif products:
            arch = (
                odt_build.Architecture.X64
                if getattr(args, "odt_arch", "64") == "64"
                else odt_build.Architecture.X86
            )
            channel = odt_build.UpdateChannel.CURRENT
            channel_arg = getattr(args, "odt_channel", None)
            if channel_arg:
                for ch in odt_build.UpdateChannel:
                    if ch.name.lower() == channel_arg.lower().replace("-", "_"):
                        channel = ch
                        break
                    if ch.value.lower() == channel_arg.lower():
                        channel = ch
                        break

            product_configs = [
                odt_build.ProductConfig(pid, languages=languages) for pid in products
            ]
            config = odt_build.ODTConfig(
                products=product_configs,
                architecture=arch,
                channel=channel,
            )
        else:
            # Default to M365 ProPlus
            config = odt_build.ODTConfig.from_preset("365-proplus-x64", languages)

        xml_output = odt_build.build_download_xml(config, download_path)

        output_path = getattr(args, "odt_output", None)
        if output_path:
            pathlib.Path(output_path).write_text(xml_output, encoding="utf-8")
            print(f"ODT download configuration written to: {output_path}")
        else:
            print(xml_output)

    except ValueError as e:
        print(f"Error: {e}", file=sys.stderr)

    return True


def _handle_repair_mode(
    args: argparse.Namespace,
    mode: str,
    human_log: logging.Logger,
    machine_log: logging.Logger,
) -> int:
    """!
    @brief Handle Office repair operations.
    @details Dispatches to quick/full repair or custom XML configuration based
    on the mode string and command-line arguments.
    @param args Parsed command-line arguments.
    @param mode Mode string (repair:quick, repair:full, or repair:config).
    @param human_log Human-readable logger.
    @param machine_log Machine-readable (JSONL) logger.
    @returns Exit code (0 for success, non-zero for failure).
    """
    _progress("-" * 60)
    _progress("Office Click-to-Run Repair Mode")
    _progress("-" * 60)

    dry_run = bool(getattr(args, "dry_run", False))
    culture = getattr(args, "repair_culture", "en-us")
    platform = getattr(args, "repair_platform", None)
    silent = not getattr(args, "repair_visible", False)

    # Check if C2R Office is installed
    _progress("Checking for Click-to-Run installation...", newline=False)
    if not repair.is_c2r_office_installed():
        _progress_fail("not found")
        human_log.error("No Click-to-Run Office installation detected")
        machine_log.info(
            "repair.error",
            extra={"event": "repair.error", "error": "c2r_not_installed"},
        )
        print("\nError: No Click-to-Run Office installation found.")
        print("This repair option only works with Office C2R installations.")
        print("\nFor MSI-based installations, use the standard Windows repair:")
        print("  Control Panel > Programs > Programs and Features > [Office] > Change > Repair")
        return 1
    _progress_ok()

    # Get installed Office info
    _progress("Gathering installation details...", newline=False)
    c2r_info = repair.get_installed_c2r_info()
    _progress_ok()
    _progress(f"  Version: {c2r_info.get('version', 'unknown')}", indent=1)
    _progress(f"  Platform: {c2r_info.get('platform', 'unknown')}", indent=1)
    _progress(f"  Culture: {c2r_info.get('culture', 'unknown')}", indent=1)
    _progress(f"  Products: {c2r_info.get('product_ids', 'unknown')}", indent=1)

    # Handle custom XML configuration
    if mode == "repair:config":
        config_path = pathlib.Path(getattr(args, "repair_config", ""))
        if not config_path.exists():
            _progress(f"Configuration file not found: {config_path}", newline=False)
            _progress_fail()
            return 1
        _progress(f"Using custom configuration: {config_path}")
        result = repair.reconfigure_office(
            config_path,
            dry_run=dry_run,
        )
        if result.returncode == 0 or result.skipped:
            _progress("Reconfiguration completed successfully", newline=False)
            _progress_ok()
            return 0
        _progress(f"Reconfiguration failed: {result.stderr or result.error}", newline=False)
        _progress_fail()
        return 1

    # Determine repair type
    repair_type_str = mode.split(":")[-1]  # quick or full
    _progress(f"Repair type: {repair_type_str.upper()}")

    if repair_type_str == "full":
        _progress("\n⚠️  WARNING: Full Online Repair may reinstall excluded applications!")
        _progress("    This operation requires internet connectivity and may take 30-60 minutes.\n")
        config = repair.RepairConfig.full_repair(
            platform=platform,
            culture=culture,
            silent=silent,
        )
    else:
        _progress("Quick Repair runs locally and typically completes in 5-15 minutes.")
        config = repair.RepairConfig.quick_repair(
            platform=platform,
            culture=culture,
            silent=silent,
        )

    # Confirm with user unless in auto mode
    if not dry_run and not getattr(args, "force", False):
        _progress("Confirm repair operation?", newline=False)
        proceed = confirm.request_scrub_confirmation(dry_run=dry_run, force=False)
        if not proceed:
            _progress_skip("cancelled by user")
            human_log.info("Repair cancelled by user")
            return 0
        _progress_ok()

    # Execute repair
    _progress("=" * 60)
    _progress(f"Executing {repair_type_str.upper()} repair...")
    _progress("=" * 60)

    repair_result = repair.run_repair(config, dry_run=dry_run)

    _progress("=" * 60)
    if repair_result.success or repair_result.skipped:
        _progress(f"Repair completed: {repair_result.summary}")
        _progress("=" * 60)
        if repair_result.skipped:
            print(f"\n[DRY-RUN] {repair_result.summary}")
        else:
            print(f"\n✓ {repair_result.summary}")
            print("\nNote: A system restart may be required to complete the repair.")
        return 0
    else:
        _progress(f"Repair failed: {repair_result.summary}")
        _progress("=" * 60)
        print(f"\n✗ {repair_result.summary}")
        if repair_result.stderr:
            print(f"\nError details:\n{repair_result.stderr}")
        return 1


def _handle_oem_config_mode(
    args: argparse.Namespace,
    mode: str,
    human_log: logging.Logger,
    machine_log: logging.Logger,
) -> int:
    """!
    @brief Handle OEM configuration execution.
    @details Executes a bundled or custom XML configuration using ODT setup.exe.
    @param args Parsed command-line arguments.
    @param mode Mode string (oem-config:<preset-name>).
    @param human_log Human-readable logger.
    @param machine_log Machine-readable (JSONL) logger.
    @returns Exit code (0 for success, non-zero for failure).
    """
    _progress("-" * 60)
    _progress("OEM Configuration Mode")
    _progress("-" * 60)

    dry_run = bool(getattr(args, "dry_run", False))
    preset_name = mode.split(":", 1)[-1] if ":" in mode else getattr(args, "oem_config", "")

    # List available presets if none specified
    if not preset_name:
        _progress("Available OEM configuration presets:")
        for name, filename, exists in repair.list_oem_configs():
            status = "✓" if exists else "✗ (missing)"
            _progress(f"  {name}: {filename} {status}", indent=1)
        return 0

    # Resolve the config path
    config_path = repair.get_oem_config_path(preset_name)
    if config_path is None:
        _progress(f"OEM config not found: {preset_name}", newline=False)
        _progress_fail()
        human_log.error(f"OEM config preset not found: {preset_name}")
        machine_log.info(
            "oem_config.error",
            extra={"event": "oem_config.error", "preset": preset_name, "error": "not_found"},
        )
        _progress("\nAvailable presets:")
        for name, _filename, exists in repair.list_oem_configs():
            if exists:
                _progress(f"  {name}", indent=1)
        return 1

    _progress(f"Preset: {preset_name}")
    _progress(f"Config file: {config_path}")

    # Check for ODT setup.exe
    setup_exe = repair.find_odt_setup_exe()
    if setup_exe is None:
        _progress("ODT setup.exe not found", newline=False)
        _progress_fail()
        human_log.error("ODT setup.exe not found")
        machine_log.info(
            "oem_config.error",
            extra={"event": "oem_config.error", "error": "setup_not_found"},
        )
        print("\nError: ODT setup.exe not found.")
        print("Please ensure setup.exe is in the oem/ folder or download it from:")
        print("  https://www.microsoft.com/en-us/download/details.aspx?id=49117")
        return 1

    _progress(f"Setup.exe: {setup_exe}")

    # Warn about destructive operations
    if preset_name in ("full-removal",):
        _progress("\n⚠️  WARNING: This will REMOVE all Office installations!")
        _progress("    This action cannot be undone.\n")
    elif "repair" in preset_name.lower():
        _progress("\nNote: Repair operations may take 5-60 minutes depending on type.\n")

    # Confirm with user unless forced
    if not dry_run and not getattr(args, "force", False):
        _progress("Confirm operation?", newline=False)
        proceed = confirm.request_scrub_confirmation(dry_run=dry_run, force=False)
        if not proceed:
            _progress_skip("cancelled by user")
            human_log.info("OEM config cancelled by user")
            return 0
        _progress_ok()

    # Execute
    _progress("=" * 60)
    _progress(f"Executing configuration: {preset_name}")
    _progress("=" * 60)

    result = repair.run_oem_config(
        preset_name,
        dry_run=dry_run,
    )

    _progress("=" * 60)
    if result.returncode == 0 or result.skipped:
        _progress("Configuration completed successfully")
        _progress("=" * 60)
        if result.skipped:
            print(f"\n[DRY-RUN] Would execute: setup.exe /configure {config_path}")
        else:
            print(f"\n✓ Configuration '{preset_name}' applied successfully.")
            print("\nNote: A system restart may be required to complete changes.")
        return 0
    else:
        _progress(f"Configuration failed: {result.stderr or result.error}")
        _progress("=" * 60)
        print(f"\n✗ Configuration '{preset_name}' failed.")
        if result.stderr:
            print(f"\nError details:\n{result.stderr}")
        return 1


def _determine_mode(args: argparse.Namespace) -> str:
    """!
    @brief Map parsed arguments to a simple textual mode identifier.
    """

    if getattr(args, "auto_all", False):
        return "auto-all"
    if getattr(args, "target", None):
        return f"target:{args.target}"
    if getattr(args, "diagnose", False):
        return "diagnose"
    if getattr(args, "cleanup_only", False):
        return "cleanup-only"
    if getattr(args, "repair", None):
        return f"repair:{args.repair}"
    if getattr(args, "repair_config", None):
        return "repair:config"
    if getattr(args, "oem_config", None):
        return f"oem-config:{args.oem_config}"
    return "interactive"


def _should_use_tui(args: argparse.Namespace) -> bool:
    """!
    @brief Determine whether the TUI should be launched automatically.
    @details The logic prefers the richer interface when the output stream
    supports ANSI escape codes and the caller did not explicitly disable color.
    """

    if getattr(args, "no_color", False):
        return False
    if getattr(sys.stdout, "isatty", None) and sys.stdout.isatty():
        return bool(os.environ.get("WT_SESSION") or os.environ.get("TERM", "").lower() != "dumb")
    return False


def _build_app_state(
    args: argparse.Namespace, human_log: logging.Logger, machine_log: logging.Logger
) -> AppState:
    """!
    @brief Assemble the dependency dictionary consumed by CLI/TUI front-ends.
    @details The mapping exposes callables for detection, planning, and
    execution so interactive interfaces can drive the same back-end flows as
    the non-interactive CLI code path.
    """

    ui_events = new_event_queue()

    def emit_event(event: str, *, message: str | None = None, **payload: object) -> None:
        """!
        @brief Publish progress events for interactive front-ends.
        @details Events are queued for consumption by the CLI/TUI layers so they
        can surface background activity without polling implementation details.
        """

        record: dict[str, object] = {"event": event}
        if message is not None:
            record["message"] = message
        if payload:
            record["data"] = dict(payload)
        ui_events.append(record)

    logging_ext.register_ui_event_sink(emitter=emit_event, queue=ui_events)

    def detector() -> dict[str, object]:
        logdir_path = pathlib.Path(
            getattr(args, "logdir", _resolve_log_directory(None))
        ).expanduser()
        return _run_detection(
            machine_log,
            logdir_path,
            limited_user=bool(getattr(args, "limited_user", False)),
        )

    def planner(
        inventory: Mapping[str, object], overrides: Mapping[str, object] | None = None
    ) -> list[dict[str, object]]:
        mode = _determine_mode(args)
        merged = dict(_collect_plan_options(args, mode))
        if overrides:
            merged.update({key: overrides[key] for key in overrides})
        generated_plan = plan_module.build_plan(dict(inventory), merged)
        safety.perform_preflight_checks(generated_plan)
        return generated_plan

    def executor(
        plan_data: list[dict[str, object]], overrides: Mapping[str, object] | None = None
    ) -> bool:
        dry_run = bool(getattr(args, "dry_run", False))
        if overrides and "dry_run" in overrides:
            dry_run = bool(overrides["dry_run"])

        mode_override = _determine_mode(args)
        if overrides and overrides.get("mode"):
            mode_override = str(overrides["mode"])

        inventory_override = overrides.get("inventory") if overrides else None
        _handle_plan_artifacts(args, plan_data, inventory_override, human_log, mode_override)

        if mode_override == "diagnose":
            human_log.info("Diagnostics complete; plan written and no actions executed.")
            return True

        guard_options = dict(_collect_plan_options(args, mode_override))
        guard_options["dry_run"] = dry_run
        if overrides:
            guard_options.update({key: overrides[key] for key in overrides})

        force_override = bool(getattr(args, "force", False))
        if overrides and "force" in overrides:
            force_override = bool(overrides["force"])

        confirmed = bool(overrides.get("confirmed")) if overrides else False
        if not confirmed:
            proceed = confirm.request_scrub_confirmation(
                dry_run=dry_run,
                force=force_override,
                input_func=overrides.get("input_func") if overrides else None,
                interactive=overrides.get("interactive") if overrides else None,
            )
            if not proceed:
                if human_log:
                    human_log.info("Scrub cancelled by user confirmation prompt.")
                if machine_log:
                    machine_log.info(
                        "scrub.cancelled",
                        extra={
                            "event": "scrub.cancelled",
                            "data": {"reason": "user_declined", "dry_run": dry_run},
                        },
                    )
                return False

        _enforce_runtime_guards(guard_options, dry_run=dry_run)
        scrub.execute_plan(plan_data, dry_run=dry_run, start_time=_MAIN_START_TIME)
        return True

    app_state: AppState = {
        "args": args,
        "human_logger": human_log,
        "machine_logger": machine_log,
        "detector": detector,
        "planner": planner,
        "executor": executor,
        "event_queue": ui_events,
        "emit_event": emit_event,
        "confirm": confirm.request_scrub_confirmation,
    }
    return app_state


def _collect_plan_options(args: argparse.Namespace, mode: str) -> dict[str, object]:
    """!
    @brief Translate parsed CLI arguments into planning options.
    @details Options are resolved with the following precedence (highest first):
    1. CLI arguments explicitly specified
    2. JSON config file values (if --config provided)
    3. Built-in defaults
    """

    # Load config file if specified
    config = _load_config_file(getattr(args, "config", None))

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

    # Resolve max_passes: --registry-only or --skip-uninstall (=0) > --passes > --max-passes > config > default (1)
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


def _run_detection(
    machine_log: logging.Logger,
    log_directory: pathlib.Path | str | None = None,
    *,
    limited_user: bool | None = None,
) -> dict[str, object]:
    """!
    @brief Execute inventory gathering, persist artifacts, and emit telemetry.
    """
    _progress("Starting inventory scan...", indent=1)

    if limited_user:
        machine_log.info("Detection requested under limited user token.")
        _progress("Running under limited user token", indent=2)

    # Use detailed progress callback for inventory gathering.
    # Since detect.py calls this from multiple threads concurrently, we use
    # complete single-line output for thread-safety (no pending line pattern).
    def progress_callback(phase: str, status: str = "start") -> None:
        prefix = "      "  # indent=3 equivalent
        timestamp = f"[{_get_elapsed_secs():12.6f}]"
        spinner.pause_for_output()
        try:
            if status == "start":
                # Print complete line with "..." to indicate in-progress
                with _PROGRESS_LOCK:
                    print(f"{timestamp} {prefix}{phase}...", flush=True)
            elif status == "ok":
                with _PROGRESS_LOCK:
                    print(
                        f"{timestamp} {prefix}{phase} [  \033[32mOK\033[0m  ]",
                        flush=True,
                    )
            elif status == "skip":
                with _PROGRESS_LOCK:
                    print(
                        f"{timestamp} {prefix}{phase} [ \033[33mSKIP\033[0m ]",
                        flush=True,
                    )
            elif status == "fail":
                with _PROGRESS_LOCK:
                    print(
                        f"{timestamp} {prefix}{phase} [\033[31mFAILED\033[0m]",
                        flush=True,
                    )
        finally:
            spinner.resume_after_output()

    _progress("Gathering Office inventory...", indent=2)
    try:
        if limited_user:
            inventory = detect.gather_office_inventory(
                limited_user=True, progress_callback=progress_callback
            )
        else:
            inventory = detect.gather_office_inventory(progress_callback=progress_callback)
        _progress("Inventory collection complete", indent=2, newline=False)
        _progress_ok()
    except KeyboardInterrupt:
        print(flush=True)  # Newline after partial output
        _progress("Inventory collection interrupted", indent=2, newline=False)
        _progress_skip("user cancelled")
        raise

    if log_directory is None:
        logdir_path = _resolve_log_directory(None)
    else:
        logdir_path = pathlib.Path(log_directory).expanduser()

    _progress(f"Log directory: {logdir_path}", indent=2)

    inventory_path: pathlib.Path | None = None
    try:
        logdir_path.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%d-%H%M%S")
        inventory_path = logdir_path / f"inventory-{timestamp}.json"
        _progress(f"Writing inventory to {inventory_path.name}...", indent=2, newline=False)
        inventory_path.write_text(
            json.dumps(inventory, indent=2, sort_keys=True, default=str),
            encoding="utf-8",
        )
        _progress_ok()
    except OSError as exc:
        _progress_fail(str(exc))
        machine_log.warning(
            "inventory_write_failed",
            extra={
                "event": "inventory_write_failed",
                "error": repr(exc),
                "logdir": str(logdir_path),
            },
        )

    # Log inventory summary
    _progress("Inventory summary:", indent=2)
    for key, value in inventory.items():
        count = len(value) if hasattr(value, "__len__") else len(list(value))
        if count > 0:
            _progress(f"{key}: {count} items", indent=3)

    machine_log.info(
        "inventory",
        extra={
            "event": "inventory",
            "counts": {
                key: len(value) if hasattr(value, "__len__") else len(list(value))
                for key, value in inventory.items()
            },
            **({"artifact": str(inventory_path)} if inventory_path is not None else {}),
        },
    )
    return inventory


def _handle_plan_artifacts(
    args: argparse.Namespace,
    plan_data: Iterable[Mapping[str, object]],
    inventory: Mapping[str, object] | None,
    human_log: logging.Logger,
    mode: str,
) -> None:
    """!
    @brief Persist plan diagnostics and backups as requested via CLI flags.
    """
    _progress("Processing plan artifacts...", indent=1)

    plan_steps = list(plan_data)
    logdir = pathlib.Path(getattr(args, "logdir", _resolve_log_directory(None))).expanduser()
    logdir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%d-%H%M%S")
    resolved_backup: pathlib.Path | None = None

    backup_dir = getattr(args, "backup", None)
    if backup_dir:
        _progress(f"Setting up backup directory: {backup_dir}", indent=2)
        destination = pathlib.Path(backup_dir).expanduser().resolve()
        destination.mkdir(parents=True, exist_ok=True)
        if inventory is not None:
            _progress("Writing inventory to backup...", indent=3, newline=False)
            (destination / "inventory.json").write_text(
                json.dumps(inventory, indent=2, sort_keys=True),
                encoding="utf-8",
            )
            _progress_ok()
        resolved_backup = destination
    else:
        resolved_backup = logdir / f"registry-backup-{timestamp}"

    if mode == "diagnose" and inventory is not None and not backup_dir:
        inventory_path = logdir / "diagnostics-inventory.json"
        _progress(f"Writing diagnostics inventory: {inventory_path.name}", indent=2)
        inventory_path.write_text(
            json.dumps(inventory, indent=2, sort_keys=True),
            encoding="utf-8",
        )
        human_log.info("Wrote diagnostics inventory to %s", inventory_path)

    if plan_steps:
        _progress(f"Enriching {len(plan_steps)} plan steps with metadata...", indent=2)
        context_step = next(
            (step for step in plan_steps if step.get("category") == "context"), None
        )
        if context_step is not None:
            metadata = dict(context_step.get("metadata", {}))
            options = dict(metadata.get("options", {}))
            metadata["log_directory"] = str(logdir)
            options["log_directory"] = str(logdir)
            options["logdir"] = str(logdir)
            if resolved_backup is not None:
                metadata["backup_destination"] = str(resolved_backup)
                options["backup"] = str(resolved_backup)
            metadata["options"] = options
            context_step["metadata"] = metadata

        if resolved_backup is not None:
            registry_steps = sum(1 for s in plan_steps if s.get("category") == "registry-cleanup")
            if registry_steps > 0:
                _progress(f"Configuring backup for {registry_steps} registry steps", indent=2)
            for step in plan_steps:
                if step.get("category") != "registry-cleanup":
                    continue
                registry_metadata = dict(step.get("metadata", {}))
                registry_metadata.setdefault("backup_destination", str(resolved_backup))
                registry_metadata.setdefault("log_directory", str(logdir))
                step["metadata"] = registry_metadata

    serialized_plan = json.dumps(plan_steps, indent=2, sort_keys=True)
    primary_plan_path = logdir / f"plan-{timestamp}.json"
    _progress(f"Writing primary plan: {primary_plan_path.name}", indent=2)
    primary_plan_path.write_text(serialized_plan, encoding="utf-8")
    human_log.info("Wrote plan to %s", primary_plan_path)

    additional_plan_targets: list[pathlib.Path] = []
    if getattr(args, "plan", None):
        additional_plan_targets.append(pathlib.Path(args.plan).expanduser().resolve())
    elif mode == "diagnose":
        additional_plan_targets.append(logdir / "diagnostics-plan.json")

    for target in additional_plan_targets:
        if target == primary_plan_path:
            continue
        _progress(f"Writing additional plan: {target.name}", indent=2)
        target.write_text(serialized_plan, encoding="utf-8")
        human_log.info("Wrote plan to %s", target)

    if backup_dir and resolved_backup is not None:
        _progress(f"Writing plan to backup: {resolved_backup}", indent=2)
        (resolved_backup / "plan.json").write_text(serialized_plan, encoding="utf-8")
        human_log.info("Wrote backup artifacts to %s", resolved_backup)

    _progress("Artifact processing complete", indent=1)


def _enforce_runtime_guards(options: Mapping[str, object], *, dry_run: bool) -> None:
    """!
    @brief Evaluate runtime safety prerequisites prior to executing the scrubber.
    @details Gathers host telemetry and forwards it to
    :func:`safety.evaluate_runtime_environment` so operating system, process, and
    restore point guards are enforced consistently across CLI entry points.
    """
    _progress("Gathering runtime environment info...", indent=1)

    _progress("Detecting operating system...", indent=2, newline=False)
    system, release = _detect_operating_system()
    _progress_ok(f"{system} {release}")

    require_restore_point = bool(options.get("create_restore_point", False))
    restore_point_available = True
    if require_restore_point and not dry_run:
        _progress("Checking restore point availability...", indent=2, newline=False)
        restore_point_available = _restore_points_available()
        if restore_point_available:
            _progress_ok()
        else:
            _progress_fail("not available")

    _progress("Checking admin privileges...", indent=2, newline=False)
    is_admin = _current_process_is_admin()
    if is_admin:
        _progress_ok()
    else:
        _progress_fail()

    _progress("Scanning for blocking processes...", indent=2, newline=False)
    blocking = _discover_blocking_processes()
    if blocking:
        _progress_fail(f"{len(blocking)} found")
        for proc in blocking[:5]:  # Show first 5
            _progress(f"- {proc}", indent=3)
        if len(blocking) > 5:
            _progress(f"... and {len(blocking) - 5} more", indent=3)
    else:
        _progress_ok("none")

    _progress("Evaluating safety constraints...", indent=2, newline=False)
    safety.evaluate_runtime_environment(
        is_admin=is_admin,
        os_system=system,
        os_release=release,
        blocking_processes=blocking,
        dry_run=dry_run,
        require_restore_point=require_restore_point,
        restore_point_available=restore_point_available,
        force=bool(options.get("force", False)),
        allow_unsupported_windows=bool(options.get("allow_unsupported_windows", False)),
        minimum_free_space_bytes=options.get("minimum_free_space_bytes"),
        disk_usage_root=options.get("disk_usage_root")
        or options.get("free_space_root")
        or options.get("system_drive"),
    )
    _progress_ok()


def _detect_operating_system() -> tuple[str, str]:
    """!
    @brief Collect the current operating system identifier and release version.
    """

    try:
        system = platform.system()
    except Exception:
        system = ""

    release = ""
    try:
        release = platform.version()
    except Exception:
        release = ""

    if not release:
        try:
            release = platform.release()
        except Exception:
            release = ""

    return system, release


def _discover_blocking_processes() -> list[str]:
    """!
    @brief Enumerate Office-related processes that may block destructive actions.
    """

    patterns = list(constants.DEFAULT_OFFICE_PROCESSES) + list(constants.OFFICE_PROCESS_PATTERNS)
    try:
        return processes.enumerate_processes(patterns)
    except Exception:
        return []


def _current_process_is_admin() -> bool:
    """!
    @brief Determine whether the current interpreter is running with elevated privileges.
    """

    if os.name == "nt":
        try:
            shell32 = ctypes.windll.shell32
            return bool(shell32.IsUserAnAdmin())
        except Exception:
            return False

    geteuid = getattr(os, "geteuid", None)
    if callable(geteuid):
        try:
            return bool(geteuid() == 0)
        except Exception:
            return False

    return False


def _restore_points_available() -> bool:
    """!
    @brief Detect whether system restore points are currently available.
    """

    if os.name != "nt":
        return False

    script = "\n".join(
        (
            "Try {",
            "  Get-ComputerRestorePoint -ErrorAction Stop | Select-Object -First 1 | Out-String",
            "  Exit 0",
            " } Catch { Exit 1 }",
        )
    )

    command = [
        "powershell.exe",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        script,
    ]

    result = exec_utils.run_command(
        command,
        event="restore_point_probe",
        timeout=15,
    )

    return result.returncode == 0 and not result.error


if __name__ == "__main__":  # pragma: no cover - for manual execution
    sys.exit(main())
