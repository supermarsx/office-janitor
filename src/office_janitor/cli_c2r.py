"""!
@brief CLI options for the 'c2r' subcommand (Click-to-Run passthrough).
@details Defines C2R direct control arguments and help text.
"""

from __future__ import annotations

import argparse

__all__ = [
    "C2R_EPILOG",
    "add_c2r_subcommand_options",
]

C2R_EPILOG = """\
C2R ACTIONS:
  --repair quick        Quick local repair (5-15 min)
  --repair full         Full online repair (30-60 min)
  --remove              Uninstall via OfficeClickToRun.exe
  --update              Trigger update check
  --rollback            Rollback to previous version

EXAMPLES:
  office-janitor c2r --repair quick          # Quick repair
  office-janitor c2r --repair full           # Full online repair
  office-janitor c2r --remove                # Uninstall C2R Office
  office-janitor c2r --update                # Check for updates
  office-janitor c2r --scenario install --config config.xml
  office-janitor c2r --args "/configure config.xml"
"""


def add_c2r_subcommand_options(parser: argparse.ArgumentParser) -> None:
    """!
    @brief Add options specific to the 'c2r' subcommand.
    @param parser The subparser to add arguments to.
    """
    # Action selection
    action_opts = parser.add_argument_group("C2R Action")
    action_mode = action_opts.add_mutually_exclusive_group()
    action_mode.add_argument(
        "-r",
        "--repair",
        choices=["quick", "full"],
        metavar="TYPE",
        help="Repair C2R Office (quick=local, full=online).",
    )
    action_mode.add_argument(
        "-R",
        "--remove",
        action="store_true",
        dest="c2r_remove",
        help="Uninstall C2R Office via OfficeClickToRun.exe.",
    )
    action_mode.add_argument(
        "-U",
        "--update",
        action="store_true",
        dest="c2r_update",
        help="Trigger Office update check.",
    )
    action_mode.add_argument(
        "--rollback",
        action="store_true",
        dest="c2r_rollback",
        help="Rollback to previous Office version.",
    )
    action_mode.add_argument(
        "-S",
        "--scenario",
        metavar="NAME",
        dest="c2r_scenario",
        help="Run a specific C2R scenario (install, repair, uninstall, etc.).",
    )

    # Repair options
    repair_opts = parser.add_argument_group("Repair Options")
    repair_opts.add_argument(
        "-l",
        "--culture",
        metavar="LANG",
        default="en-us",
        dest="c2r_culture",
        help="Language/culture code for repair (default: en-us).",
    )
    repair_opts.add_argument(
        "-p",
        "--platform",
        choices=["x86", "x64"],
        metavar="ARCH",
        dest="c2r_platform",
        help="Architecture (auto-detected if not specified).",
    )
    repair_opts.add_argument(
        "-V",
        "--visible",
        action="store_true",
        dest="c2r_visible",
        help="Show repair UI instead of running silently.",
    )
    repair_opts.add_argument(
        "-t",
        "--timeout",
        type=int,
        default=3600,
        metavar="SEC",
        dest="c2r_timeout",
        help="Timeout in seconds (default: 3600).",
    )

    # Advanced options
    advanced_opts = parser.add_argument_group("Advanced Options")
    advanced_opts.add_argument(
        "-c",
        "--config",
        metavar="XML",
        dest="c2r_config",
        help="XML configuration file to use with scenario.",
    )
    advanced_opts.add_argument(
        "--args",
        metavar="ARGS",
        dest="c2r_args",
        help="Raw arguments to pass to OfficeClickToRun.exe.",
    )
    advanced_opts.add_argument(
        "--productstoremove",
        metavar="IDS",
        dest="c2r_products_to_remove",
        help="Comma-separated product IDs to remove.",
    )
    advanced_opts.add_argument(
        "-F",
        "--force-app-shutdown",
        action="store_true",
        dest="c2r_force_shutdown",
        help="Force close running Office applications.",
    )
    advanced_opts.add_argument(
        "--display-level",
        choices=["None", "Full"],
        default="None",
        metavar="LEVEL",
        dest="c2r_display_level",
        help="Display level for UI (None=silent, Full=visible).",
    )
