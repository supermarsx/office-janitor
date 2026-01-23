"""!
@brief CLI options for the 'repair' subcommand.
@details Defines repair-specific arguments and help text for Office Janitor.
"""

from __future__ import annotations

import argparse

__all__ = [
    "REPAIR_EPILOG",
    "add_repair_subcommand_options",
]

REPAIR_EPILOG = """\
REPAIR MODES:
  (default)    Auto-detect and repair all Office installations
  --quick      Quick local repair - fixes common issues (5-15 min)
  --full       Full online repair - redownloads components (30-60 min)
  --odt        Repair using Office Deployment Tool configuration
  --c2r        Repair using OfficeClickToRun.exe directly

EXAMPLES:
  office-janitor repair                      # Auto-repair all detected
  office-janitor repair --quick              # Quick local repair
  office-janitor repair --full --visible     # Full repair with UI
  office-janitor repair --dry-run            # Preview repair operations
"""


def add_repair_subcommand_options(parser: argparse.ArgumentParser) -> None:
    """!
    @brief Add options specific to the 'repair' subcommand.
    @param parser The subparser to add arguments to.
    """
    repair_mode = parser.add_argument_group("Repair Mode")
    repair_type = repair_mode.add_mutually_exclusive_group()
    repair_type.add_argument(
        "-q",
        "--quick",
        action="store_const",
        const="quick",
        dest="repair_type",
        help="Quick local repair (5-15 min, no internet required).",
    )
    repair_type.add_argument(
        "-F",
        "--full",
        action="store_const",
        const="full",
        dest="repair_type",
        help="Full online repair (30-60 min, downloads components).",
    )
    repair_type.add_argument(
        "--odt",
        action="store_true",
        dest="repair_odt",
        help="Repair using ODT/setup.exe configuration method.",
    )
    repair_type.add_argument(
        "--c2r",
        action="store_true",
        dest="repair_c2r",
        help="Repair using OfficeClickToRun.exe directly.",
    )

    repair_opts = parser.add_argument_group("Repair Options")
    repair_opts.add_argument(
        "-c",
        "--culture",
        metavar="LANG",
        default="en-us",
        dest="repair_culture",
        help="Language/culture code for repair (default: en-us).",
    )
    repair_opts.add_argument(
        "-p",
        "--platform",
        choices=["x86", "x64"],
        metavar="ARCH",
        dest="repair_platform",
        help="Architecture for repair (auto-detected if not specified).",
    )
    repair_opts.add_argument(
        "-V",
        "--visible",
        action="store_true",
        dest="repair_visible",
        help="Show repair UI instead of running silently.",
    )
    repair_opts.add_argument(
        "-t",
        "--timeout",
        type=int,
        default=3600,
        metavar="SEC",
        dest="repair_timeout",
        help="Timeout for repair operations in seconds (default: 3600).",
    )
    repair_opts.add_argument(
        "-a",
        "--all-products",
        action="store_true",
        dest="repair_all_products",
        help="Repair all detected Office products.",
    )
    repair_opts.add_argument(
        "-P",
        "--preset",
        metavar="NAME",
        dest="repair_preset",
        help="Use a specific repair preset (quick-repair, full-repair, etc.).",
    )
