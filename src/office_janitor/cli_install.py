"""!
@brief CLI options for the 'install' subcommand.
@details Defines installation-specific arguments and help text for Office Janitor.
"""

from __future__ import annotations

import argparse

__all__ = [
    "INSTALL_EPILOG",
    "add_install_subcommand_options",
]

INSTALL_EPILOG = """\
INSTALLATION PRESETS:
  365-proplus-x64         Microsoft 365 Apps for enterprise (64-bit)
  365-business-x64        Microsoft 365 Apps for business (64-bit)
  office2024-x64          Office LTSC 2024 Professional Plus (64-bit)
  office2021-x64          Office LTSC 2021 Professional Plus (64-bit)
  ltsc2024-full-x64       Office 2024 + Visio + Project (64-bit)
  ltsc2024-full-x64-clean Office 2024 + Visio + Project (no bloat)

EXAMPLES:
  office-janitor install --preset 365-proplus-x64
  office-janitor install --preset ltsc2024-full-x64 --language pt-pt
  office-janitor install --goobler              # LTSC 2024 + Visio + Project
  office-janitor install --build config.xml    # Generate custom ODT config
"""


def add_install_subcommand_options(parser: argparse.ArgumentParser) -> None:
    """!
    @brief Add options specific to the 'install' subcommand.
    @param parser The subparser to add arguments to.
    """
    install_opts = parser.add_argument_group("Installation Options")
    install_opts.add_argument(
        "-p",
        "--preset",
        metavar="NAME",
        dest="odt_preset",
        help="Use a predefined ODT installation preset (use --list-presets to see options).",
    )
    install_opts.add_argument(
        "-P",
        "--product",
        metavar="ID",
        action="append",
        dest="odt_products",
        help="Product ID to include in installation (repeatable).",
    )
    install_opts.add_argument(
        "-l",
        "--language",
        metavar="CODE",
        action="append",
        dest="odt_languages",
        help="Language code for installation (repeatable, default: en-us).",
    )
    install_opts.add_argument(
        "-a",
        "--arch",
        choices=["32", "64"],
        default="64",
        dest="odt_arch",
        metavar="BITS",
        help="Architecture for installation (default: 64).",
    )
    install_opts.add_argument(
        "-C",
        "--channel",
        metavar="CHANNEL",
        dest="odt_channel",
        help="Update channel for installation.",
    )
    install_opts.add_argument(
        "-s",
        "--shared-computer",
        action="store_true",
        dest="odt_shared_computer",
        help="Enable shared computer licensing.",
    )
    install_opts.add_argument(
        "--remove-msi",
        action="store_true",
        dest="odt_remove_msi",
        help="Remove existing MSI Office before installation.",
    )
    install_opts.add_argument(
        "-x",
        "--exclude-app",
        metavar="APP",
        action="append",
        dest="odt_exclude_apps",
        help="App to exclude from installation (repeatable).",
    )
    install_opts.add_argument(
        "--include-visio",
        action="store_true",
        dest="odt_include_visio",
        help="Include Visio Professional.",
    )
    install_opts.add_argument(
        "--include-project",
        action="store_true",
        dest="odt_include_project",
        help="Include Project Professional.",
    )

    # Build/generate options
    build_opts = parser.add_argument_group("Build Options")
    build_opts.add_argument(
        "-b",
        "--build",
        metavar="FILE",
        dest="odt_output",
        help="Generate an ODT XML configuration file instead of installing.",
    )
    build_opts.add_argument(
        "-d",
        "--download",
        metavar="PATH",
        dest="odt_download",
        help="Generate a download XML with the specified local path.",
    )

    # List options
    list_opts = parser.add_argument_group("List Options")
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

    # Author aliases
    alias_opts = parser.add_argument_group("Quick Install Aliases")
    alias_opts.add_argument(
        "--goobler",
        action="store_true",
        help="Install LTSC 2024 + Visio + Project (clean) with pt-pt and en-us.",
    )
    alias_opts.add_argument(
        "--pupa",
        action="store_true",
        help="Install LTSC 2024 ProPlus only (clean) with pt-pt and en-us.",
    )
