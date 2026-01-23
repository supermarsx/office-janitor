"""!
@brief CLI options for the 'diagnose' subcommand.
@details Defines diagnostic-specific arguments and help text for Office Janitor.
"""

from __future__ import annotations

import argparse

__all__ = [
    "DIAGNOSE_EPILOG",
    "add_diagnose_subcommand_options",
]

DIAGNOSE_EPILOG = """\
OUTPUT FORMATS:
  --plan FILE           Save detailed action plan as JSON
  --json                Output structured events to stdout
  --verbose             Increase detail level (-v, -vv, -vvv)

LEGACY MODE FLAGS:
  --diagnose            Emit inventory and plan without making changes
  --target VER          Target a specific Office version for diagnostics

EXAMPLES:
  office-janitor diagnose                    # Show detected Office installations
  office-janitor diagnose --plan report.json # Save plan to file
  office-janitor diagnose --json             # Machine-readable output
  office-janitor diagnose -vvv               # Maximum verbosity
  office-janitor diagnose --target 2019      # Focus on specific version
"""


def add_diagnose_subcommand_options(parser: argparse.ArgumentParser) -> None:
    """!
    @brief Add options specific to the 'diagnose' subcommand.
    @param parser The subparser to add arguments to.
    """
    # Legacy mode flags (moved from top-level)
    legacy_opts = parser.add_argument_group("Legacy Mode Flags")
    legacy_opts.add_argument(
        "--diagnose",
        action="store_true",
        dest="legacy_diagnose",
        help="Emit inventory and plan without making changes (implicit for this subcommand).",
    )
    legacy_opts.add_argument(
        "--target",
        metavar="VER",
        help="Target a specific Office version (2003-2024/365).",
    )

    diag_opts = parser.add_argument_group("Diagnostic Options")
    diag_opts.add_argument(
        "-i",
        "--inventory",
        metavar="FILE",
        help="Save inventory data to a JSON file.",
    )
    diag_opts.add_argument(
        "-H",
        "--check-health",
        action="store_true",
        help="Perform health checks on detected installations.",
    )
    diag_opts.add_argument(
        "-L",
        "--check-licenses",
        action="store_true",
        help="Check license status of detected products.",
    )
    diag_opts.add_argument(
        "-U",
        "--check-updates",
        action="store_true",
        help="Check for available updates.",
    )
