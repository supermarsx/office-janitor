"""!
@brief CLI options for the 'license' subcommand.
@details Defines Office licensing management arguments and help text.
"""

from __future__ import annotations

import argparse

__all__ = [
    "LICENSE_EPILOG",
    "add_license_subcommand_options",
]

LICENSE_EPILOG = """\
LICENSE ACTIONS:
  --status              Show current license status for all products
  --clean               Remove all Office license tokens
  --clean-spp           Clean Software Protection Platform tokens only
  --clean-ospp          Clean Office Software Protection Platform tokens only
  --clean-vnext         Clean vNext/device-based licensing only

EXAMPLES:
  office-janitor license --status            # Show license status
  office-janitor license --clean             # Remove all license tokens
  office-janitor license --clean-spp         # SPP tokens only
  office-janitor license --clean-ospp        # OSPP tokens only
  office-janitor license --backup backup/    # Backup before cleaning
"""


def add_license_subcommand_options(parser: argparse.ArgumentParser) -> None:
    """!
    @brief Add options specific to the 'license' subcommand.
    @param parser The subparser to add arguments to.
    """
    # Action selection
    action_opts = parser.add_argument_group("License Action")
    action_mode = action_opts.add_mutually_exclusive_group()
    action_mode.add_argument(
        "-s",
        "--status",
        action="store_true",
        dest="license_status",
        help="Show current license status for all Office products.",
    )
    action_mode.add_argument(
        "-c",
        "--clean",
        action="store_true",
        dest="license_clean",
        help="Remove all Office license tokens.",
    )
    action_mode.add_argument(
        "--clean-spp",
        action="store_true",
        dest="license_clean_spp",
        help="Clean Software Protection Platform (SPP) tokens only.",
    )
    action_mode.add_argument(
        "--clean-ospp",
        action="store_true",
        dest="license_clean_ospp",
        help="Clean Office Software Protection Platform (OSPP) tokens only.",
    )
    action_mode.add_argument(
        "--clean-vnext",
        action="store_true",
        dest="license_clean_vnext",
        help="Clean vNext/device-based licensing cache only.",
    )

    # Options
    opts = parser.add_argument_group("Options")
    opts.add_argument(
        "-b",
        "--backup",
        metavar="DIR",
        dest="license_backup",
        help="Backup license data before cleaning.",
    )
    opts.add_argument(
        "--keep-kms",
        action="store_true",
        dest="license_keep_kms",
        help="Preserve KMS activation when cleaning.",
    )
    opts.add_argument(
        "--keep-mak",
        action="store_true",
        dest="license_keep_mak",
        help="Preserve MAK activation when cleaning.",
    )
    opts.add_argument(
        "-A",
        "--all-users",
        action="store_true",
        dest="license_all_users",
        help="Clean licenses for all user profiles.",
    )

    # Query options
    query_opts = parser.add_argument_group("Query Options")
    query_opts.add_argument(
        "--list-products",
        action="store_true",
        dest="license_list_products",
        help="List all licensed Office products.",
    )
    query_opts.add_argument(
        "--list-tokens",
        action="store_true",
        dest="license_list_tokens",
        help="List all license tokens (detailed).",
    )
    query_opts.add_argument(
        "--export",
        metavar="FILE",
        dest="license_export",
        help="Export license information to JSON file.",
    )
