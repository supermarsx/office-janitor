"""!
@brief CLI options for the 'odt' subcommand (ODT builder/generator).
@details Defines ODT XML configuration generation arguments and help text.
"""

from __future__ import annotations

import argparse

__all__ = [
    "ODT_EPILOG",
    "add_odt_subcommand_options",
]

ODT_EPILOG = """\
OUTPUT MODES:
  --output FILE         Write generated XML to file
  --stdout              Print generated XML to stdout
  --install             Generate and immediately run installation

EXAMPLES:
  office-janitor odt --preset 365-proplus-x64 --output config.xml
  office-janitor odt --product O365ProPlusRetail --language en-us --stdout
  office-janitor odt --preset ltsc2024-full-x64 --install
  office-janitor odt --removal --output remove.xml
  office-janitor odt --list-presets
  office-janitor odt --list-products
"""


def add_odt_subcommand_options(parser: argparse.ArgumentParser) -> None:
    """!
    @brief Add options specific to the 'odt' subcommand.
    @param parser The subparser to add arguments to.
    """
    # Preset/Product selection
    product_opts = parser.add_argument_group("Product Selection")
    product_opts.add_argument(
        "-p",
        "--preset",
        metavar="NAME",
        dest="odt_preset",
        help="Use a predefined installation preset (use --list-presets to see options).",
    )
    product_opts.add_argument(
        "-P",
        "--product",
        metavar="ID",
        action="append",
        dest="odt_products",
        help="Product ID to include (repeatable). Use --list-products to see options.",
    )
    product_opts.add_argument(
        "-l",
        "--language",
        metavar="CODE",
        action="append",
        dest="odt_languages",
        help="Language code (repeatable, default: en-us).",
    )
    product_opts.add_argument(
        "-a",
        "--arch",
        choices=["32", "64"],
        default="64",
        dest="odt_arch",
        metavar="BITS",
        help="Architecture (default: 64).",
    )
    product_opts.add_argument(
        "-C",
        "--channel",
        metavar="CHANNEL",
        dest="odt_channel",
        help="Update channel (use --list-channels to see options).",
    )

    # Configuration options
    config_opts = parser.add_argument_group("Configuration Options")
    config_opts.add_argument(
        "-x",
        "--exclude-app",
        metavar="APP",
        action="append",
        dest="odt_exclude_apps",
        help="App to exclude (repeatable): Access, Excel, OneDrive, etc.",
    )
    config_opts.add_argument(
        "--include-visio",
        action="store_true",
        dest="odt_include_visio",
        help="Include Visio Professional.",
    )
    config_opts.add_argument(
        "--include-project",
        action="store_true",
        dest="odt_include_project",
        help="Include Project Professional.",
    )
    config_opts.add_argument(
        "-s",
        "--shared-computer",
        action="store_true",
        dest="odt_shared_computer",
        help="Enable shared computer licensing.",
    )
    config_opts.add_argument(
        "--remove-msi",
        action="store_true",
        dest="odt_remove_msi",
        help="Include RemoveMSI element to remove existing MSI Office.",
    )

    # Output options
    output_opts = parser.add_argument_group("Output Options")
    output_mode = output_opts.add_mutually_exclusive_group()
    output_mode.add_argument(
        "-o",
        "--output",
        metavar="FILE",
        dest="odt_output",
        help="Write generated XML to file.",
    )
    output_mode.add_argument(
        "--stdout",
        action="store_true",
        dest="odt_stdout",
        help="Print generated XML to stdout.",
    )
    output_mode.add_argument(
        "--install",
        action="store_true",
        dest="odt_run_install",
        help="Generate config and immediately run installation.",
    )

    # Special modes
    special_opts = parser.add_argument_group("Special Modes")
    special_opts.add_argument(
        "--removal",
        action="store_true",
        dest="odt_removal",
        help="Generate a removal XML instead of installation XML.",
    )
    special_opts.add_argument(
        "--download",
        metavar="PATH",
        dest="odt_download",
        help="Generate a download XML with the specified local source path.",
    )

    # List options
    list_opts = parser.add_argument_group("List Available Options")
    list_opts.add_argument(
        "--list-presets",
        action="store_true",
        dest="odt_list_presets",
        help="List all available installation presets and exit.",
    )
    list_opts.add_argument(
        "--list-products",
        action="store_true",
        dest="odt_list_products",
        help="List all available ODT product IDs and exit.",
    )
    list_opts.add_argument(
        "--list-channels",
        action="store_true",
        dest="odt_list_channels",
        help="List all available update channels and exit.",
    )
    list_opts.add_argument(
        "--list-languages",
        action="store_true",
        dest="odt_list_languages",
        help="List all supported language codes and exit.",
    )
