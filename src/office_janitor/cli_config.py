"""!
@brief CLI options for the 'config' subcommand (configuration generator).
@details Defines configuration file generation arguments and help text.
"""

from __future__ import annotations

import argparse

__all__ = [
    "CONFIG_EPILOG",
    "add_config_subcommand_options",
]

CONFIG_EPILOG = """\
CONFIG FORMATS:
  --json                Generate JSON configuration file
  --xml                 Generate ODT XML configuration file
  --ini                 Generate INI-style configuration file

EXAMPLES:
  office-janitor config --interactive        # Interactive config wizard
  office-janitor config --json -o config.json
  office-janitor config --from-current       # Generate from current settings
  office-janitor config --template remove    # Start from template
  office-janitor config --validate config.json
"""


def add_config_subcommand_options(parser: argparse.ArgumentParser) -> None:
    """!
    @brief Add options specific to the 'config' subcommand.
    @param parser The subparser to add arguments to.
    """
    # Mode selection
    mode_opts = parser.add_argument_group("Generation Mode")
    mode_group = mode_opts.add_mutually_exclusive_group()
    mode_group.add_argument(
        "-i",
        "--interactive",
        action="store_true",
        dest="config_interactive",
        help="Run interactive configuration wizard.",
    )
    mode_group.add_argument(
        "-c",
        "--from-current",
        action="store_true",
        dest="config_from_current",
        help="Generate config from current Office installation.",
    )
    mode_group.add_argument(
        "-t",
        "--template",
        metavar="NAME",
        choices=["install", "remove", "repair", "full"],
        dest="config_template",
        help="Start from a template (install, remove, repair, full).",
    )
    mode_group.add_argument(
        "-V",
        "--validate",
        metavar="FILE",
        dest="config_validate",
        help="Validate an existing configuration file.",
    )

    # Output format
    format_opts = parser.add_argument_group("Output Format")
    format_group = format_opts.add_mutually_exclusive_group()
    format_group.add_argument(
        "--json",
        action="store_const",
        const="json",
        dest="config_format",
        help="Generate JSON configuration file.",
    )
    format_group.add_argument(
        "--xml",
        action="store_const",
        const="xml",
        dest="config_format",
        help="Generate ODT XML configuration file.",
    )
    format_group.add_argument(
        "--ini",
        action="store_const",
        const="ini",
        dest="config_format",
        help="Generate INI-style configuration file.",
    )

    # Output options
    output_opts = parser.add_argument_group("Output Options")
    output_opts.add_argument(
        "-o",
        "--output",
        metavar="FILE",
        dest="config_output",
        help="Output file path (default: stdout).",
    )
    output_opts.add_argument(
        "--stdout",
        action="store_true",
        dest="config_stdout",
        help="Print to stdout instead of file.",
    )
    output_opts.add_argument(
        "--pretty",
        action="store_true",
        dest="config_pretty",
        help="Pretty-print output with indentation.",
    )

    # Include options
    include_opts = parser.add_argument_group("Include Options")
    include_opts.add_argument(
        "--include-defaults",
        action="store_true",
        dest="config_include_defaults",
        help="Include default values in output.",
    )
    include_opts.add_argument(
        "--include-comments",
        action="store_true",
        dest="config_include_comments",
        help="Include documentation comments.",
    )
    include_opts.add_argument(
        "--include-all",
        action="store_true",
        dest="config_include_all",
        help="Include all available options.",
    )
