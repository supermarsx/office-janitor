"""!
@brief CLI options for the 'offscrub' subcommand.
@details Defines OffScrub-style removal arguments and help text.
"""

from __future__ import annotations

import argparse

__all__ = [
    "OFFSCRUB_EPILOG",
    "add_offscrub_subcommand_options",
]

OFFSCRUB_EPILOG = """\
OFFSCRUB MODES:
  --all                 Remove all detected Office products
  --msi                 Remove MSI-based Office only (2003-2016)
  --c2r                 Remove Click-to-Run Office only (2013+)
  --version VER         Target specific version (03, 07, 10, 13, 16, c2r)

LEGACY FLAGS:
  These flags mirror the original OffScrub VBS scripts for compatibility.

EXAMPLES:
  office-janitor offscrub --all              # Remove everything
  office-janitor offscrub --msi              # MSI Office only
  office-janitor offscrub --c2r              # Click-to-Run only
  office-janitor offscrub --version 16       # Office 2016 MSI only
  office-janitor offscrub --ose              # Fix OSE service only
  office-janitor offscrub --dry-run          # Preview (safe mode)
"""


def add_offscrub_subcommand_options(parser: argparse.ArgumentParser) -> None:
    """!
    @brief Add options specific to the 'offscrub' subcommand.
    @param parser The subparser to add arguments to.
    """
    # Target selection
    target_opts = parser.add_argument_group("Target Selection")
    target_mode = target_opts.add_mutually_exclusive_group()
    target_mode.add_argument(
        "-A",
        "--all",
        action="store_true",
        dest="offscrub_all",
        help="Remove all detected Office products (MSI and C2R).",
    )
    target_mode.add_argument(
        "-M",
        "--msi",
        action="store_true",
        dest="offscrub_msi",
        help="Remove MSI-based Office products only.",
    )
    target_mode.add_argument(
        "-C",
        "--c2r",
        action="store_true",
        dest="offscrub_c2r",
        help="Remove Click-to-Run Office products only.",
    )
    target_mode.add_argument(
        "-V",
        "--version",
        metavar="VER",
        dest="offscrub_version",
        choices=["03", "07", "10", "13", "15", "16", "c2r"],
        help="Target specific Office version (03, 07, 10, 13, 15, 16, c2r).",
    )

    # OffScrub flags
    offscrub_opts = parser.add_argument_group("OffScrub Options")
    offscrub_opts.add_argument(
        "--ose",
        action="store_true",
        dest="offscrub_ose",
        help="Fix OSE (Office Source Engine) service configuration.",
    )
    offscrub_opts.add_argument(
        "--offline",
        action="store_true",
        dest="offscrub_offline",
        help="Mark C2R configuration as offline mode.",
    )
    offscrub_opts.add_argument(
        "-Q",
        "--quiet",
        action="store_true",
        dest="offscrub_quiet",
        help="Reduce log verbosity (quiet mode).",
    )
    offscrub_opts.add_argument(
        "--test-rerun",
        action="store_true",
        dest="offscrub_test_rerun",
        help="Run uninstall passes twice (test rerun).",
    )
    offscrub_opts.add_argument(
        "--bypass",
        action="store_true",
        dest="offscrub_bypass",
        help="Bypass certain safety checks.",
    )
    offscrub_opts.add_argument(
        "--fast-remove",
        action="store_true",
        dest="offscrub_fast_remove",
        help="Skip verification probes for faster removal.",
    )
    offscrub_opts.add_argument(
        "--scan-components",
        action="store_true",
        dest="offscrub_scan_components",
        help="Scan Windows Installer components.",
    )
    offscrub_opts.add_argument(
        "--return-error",
        action="store_true",
        dest="offscrub_return_error",
        help="Return error codes on partial success.",
    )

    # Process control
    process_opts = parser.add_argument_group("Process Control")
    process_opts.add_argument(
        "-F",
        "--force-app-shutdown",
        action="store_true",
        help="Force close running Office applications.",
    )
    process_opts.add_argument(
        "--skip-processes",
        action="store_true",
        help="Skip terminating Office processes.",
    )
    process_opts.add_argument(
        "--skip-services",
        action="store_true",
        help="Skip stopping Office services.",
    )
