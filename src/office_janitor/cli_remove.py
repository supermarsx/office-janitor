"""!
@brief CLI options for the 'remove' subcommand.
@details Defines removal/uninstall-specific arguments and help text for Office Janitor.
"""

from __future__ import annotations

import argparse

__all__ = [
    "REMOVE_EPILOG",
    "add_remove_subcommand_options",
]

REMOVE_EPILOG = """\
SCRUB LEVELS:
  minimal      Uninstall only, minimal cleanup
  standard     Uninstall + registry/filesystem cleanup (default)
  aggressive   Deep cleanup including license artifacts
  nuclear      Complete removal of all Office traces

EXAMPLES:
  office-janitor remove                      # Remove all detected Office
  office-janitor remove --target 2019        # Target specific version
  office-janitor remove --c2r-only           # Remove Click-to-Run only
  office-janitor remove --msi-only           # Remove MSI Office only
  office-janitor remove --scrub aggressive   # Aggressive cleanup
  office-janitor remove --dry-run            # Preview (safe mode)
"""


def add_remove_subcommand_options(parser: argparse.ArgumentParser) -> None:
    """!
    @brief Add options specific to the 'remove' subcommand.
    @param parser The subparser to add arguments to.
    """
    target_opts = parser.add_argument_group("Target Selection")
    target_opts.add_argument(
        "-t",
        "--target",
        metavar="VERSION",
        help="Target specific Office version (2013, 2016, 2019, 2021, 2024, 365).",
    )
    target_opts.add_argument(
        "-M",
        "--msi-only",
        action="store_const",
        const="msi",
        dest="uninstall_method",
        help="Only uninstall MSI-based Office products.",
    )
    target_opts.add_argument(
        "-C",
        "--c2r-only",
        action="store_const",
        const="c2r",
        dest="uninstall_method",
        help="Only uninstall Click-to-Run Office products.",
    )
    target_opts.add_argument(
        "-g",
        "--product-code",
        metavar="GUID",
        action="append",
        dest="product_codes",
        help="Specific MSI product code(s) to uninstall (repeatable).",
    )
    target_opts.add_argument(
        "-r",
        "--release-id",
        metavar="ID",
        action="append",
        dest="release_ids",
        help="Specific C2R release ID(s) to uninstall (repeatable).",
    )

    scrub_opts = parser.add_argument_group("Scrub Level")
    scrub_opts.add_argument(
        "-s",
        "--scrub",
        choices=["minimal", "standard", "aggressive", "nuclear"],
        default="standard",
        dest="scrub_level",
        metavar="LEVEL",
        help="Scrub intensity level (default: standard).",
    )

    uninstall_opts = parser.add_argument_group("Uninstall Options")
    uninstall_opts.add_argument(
        "-F",
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
        "-p",
        "--passes",
        type=int,
        default=None,
        metavar="N",
        help="Number of uninstall passes (default: 1).",
    )
