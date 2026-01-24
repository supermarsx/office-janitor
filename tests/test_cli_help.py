"""!
@brief Tests for cli_help module.
@details Verifies CLI argument parsing and help text generation.
"""

from __future__ import annotations

import argparse

from office_janitor import cli_help


class TestBuildArgParser:
    """Tests for build_arg_parser function."""

    def test_parser_creation(self) -> None:
        """Parser should be created successfully."""
        parser = cli_help.build_arg_parser()
        assert isinstance(parser, argparse.ArgumentParser)
        assert parser.prog == "office-janitor"

    def test_parser_with_version_info(self) -> None:
        """Parser should accept version info."""
        version_info = {"version": "1.0.0", "build": "test"}
        parser = cli_help.build_arg_parser(version_info=version_info)
        assert parser is not None

    def test_mode_arguments_mutually_exclusive(self) -> None:
        """Legacy mode arguments are no longer mutually exclusive (hidden from usage).

        Note: These flags still work for backward compatibility but users
        should use subcommands instead (e.g., 'remove --auto-all').
        """
        parser = cli_help.build_arg_parser()
        # Should succeed with one mode
        args = parser.parse_args(["--diagnose"])
        assert args.diagnose is True

        # Legacy flags are no longer mutually exclusive to hide them from usage
        # They can be combined at the parser level but behavior is undefined
        args = parser.parse_args(["--diagnose", "--auto-all"])
        assert args.diagnose is True
        assert args.auto_all is True

    def test_auto_repair_mode(self) -> None:
        """--auto-repair should be recognized."""
        parser = cli_help.build_arg_parser()
        args = parser.parse_args(["--auto-repair"])
        assert args.auto_repair is True

    def test_repair_quick(self) -> None:
        """--repair quick should be recognized."""
        parser = cli_help.build_arg_parser()
        args = parser.parse_args(["--repair", "quick"])
        assert args.repair == "quick"

    def test_repair_full(self) -> None:
        """--repair full should be recognized."""
        parser = cli_help.build_arg_parser()
        args = parser.parse_args(["--repair", "full"])
        assert args.repair == "full"

    def test_repair_odt_flag(self) -> None:
        """--odt in repair subcommand should be recognized."""
        parser = cli_help.build_arg_parser()
        args = parser.parse_args(["repair", "--odt"])
        assert args.repair_odt is True

    def test_repair_c2r_flag(self) -> None:
        """--c2r in repair subcommand should be recognized."""
        parser = cli_help.build_arg_parser()
        args = parser.parse_args(["repair", "--c2r"])
        assert args.repair_c2r is True

    def test_repair_culture(self) -> None:
        """--culture in repair subcommand should accept language codes."""
        parser = cli_help.build_arg_parser()
        args = parser.parse_args(["repair", "--culture", "de-de"])
        assert args.repair_culture == "de-de"

    def test_repair_platform(self) -> None:
        """--platform in repair subcommand should accept x86 and x64."""
        parser = cli_help.build_arg_parser()
        args = parser.parse_args(["repair", "--platform", "x64"])
        assert args.repair_platform == "x64"

        args = parser.parse_args(["repair", "--platform", "x86"])
        assert args.repair_platform == "x86"

    def test_repair_visible(self) -> None:
        """--visible in repair subcommand should be recognized."""
        parser = cli_help.build_arg_parser()
        args = parser.parse_args(["repair", "--visible"])
        assert args.repair_visible is True

    def test_dry_run_flag(self) -> None:
        """--dry-run and -n should be recognized (global option)."""
        parser = cli_help.build_arg_parser()
        args = parser.parse_args(["--dry-run"])
        assert args.dry_run is True

        args = parser.parse_args(["-n"])
        assert args.dry_run is True

    def test_force_flag(self) -> None:
        """--force and -f in subcommand should be recognized."""
        parser = cli_help.build_arg_parser()
        args = parser.parse_args(["remove", "--force"])
        assert args.force is True

        args = parser.parse_args(["remove", "-f"])
        assert args.force is True

    def test_verbose_counting(self) -> None:
        """--verbose should count occurrences (global option)."""
        parser = cli_help.build_arg_parser()
        args = parser.parse_args(["-v"])
        assert args.verbose == 1

        args = parser.parse_args(["-vvv"])
        assert args.verbose == 3

    def test_odt_build_flags(self) -> None:
        """ODT flags in install subcommand should be recognized."""
        parser = cli_help.build_arg_parser()
        args = parser.parse_args(["install", "--build", "output.xml"])
        assert args.odt_output == "output.xml"

    def test_odt_preset(self) -> None:
        """--preset in install subcommand should accept preset names."""
        parser = cli_help.build_arg_parser()
        args = parser.parse_args(["install", "--preset", "365-proplus-x64"])
        assert args.odt_preset == "365-proplus-x64"

    def test_odt_product_list(self) -> None:
        """--product in install subcommand should accumulate products."""
        parser = cli_help.build_arg_parser()
        args = parser.parse_args(
            [
                "install",
                "--product",
                "O365ProPlusRetail",
                "--product",
                "VisioProRetail",
            ]
        )
        assert args.odt_products == ["O365ProPlusRetail", "VisioProRetail"]

    def test_odt_language_list(self) -> None:
        """--language in install subcommand should accumulate languages."""
        parser = cli_help.build_arg_parser()
        args = parser.parse_args(
            [
                "install",
                "--language",
                "en-us",
                "--language",
                "de-de",
            ]
        )
        assert args.odt_languages == ["en-us", "de-de"]


class TestHelpTextGeneration:
    """Tests for help text generation functions."""

    def test_format_repair_help(self) -> None:
        """format_repair_help should return non-empty string."""
        help_text = cli_help.format_repair_help()
        assert isinstance(help_text, str)
        assert len(help_text) > 0
        assert "AUTO-REPAIR" in help_text
        assert "QUICK REPAIR" in help_text
        assert "FULL" in help_text

    def test_format_quick_reference(self) -> None:
        """format_quick_reference should return non-empty string."""
        ref_text = cli_help.format_quick_reference()
        assert isinstance(ref_text, str)
        assert len(ref_text) > 0
        assert "REPAIR" in ref_text
        assert "REMOVE" in ref_text
        assert "DIAGNOSE" in ref_text


class TestArgumentGroups:
    """Tests for individual argument group functions."""

    def test_add_mode_arguments(self) -> None:
        """add_mode_arguments should add mode flags."""
        parser = argparse.ArgumentParser()
        cli_help.add_mode_arguments(parser)
        args = parser.parse_args(["--diagnose"])
        assert args.diagnose is True

    def test_add_core_options(self) -> None:
        """add_core_options should add core flags."""
        parser = argparse.ArgumentParser()
        cli_help.add_core_options(parser)
        args = parser.parse_args(["--force", "--dry-run"])
        assert args.force is True
        assert args.dry_run is True

    def test_add_repair_options(self) -> None:
        """add_repair_options should add repair flags."""
        parser = argparse.ArgumentParser()
        cli_help.add_repair_options(parser)
        args = parser.parse_args(["--repair-odt", "--repair-culture", "fr-fr"])
        assert args.repair_odt is True
        assert args.repair_culture == "fr-fr"

    def test_add_uninstall_options(self) -> None:
        """add_uninstall_options should add uninstall flags."""
        parser = argparse.ArgumentParser()
        cli_help.add_uninstall_options(parser)
        args = parser.parse_args(["--uninstall-method", "msi"])
        assert args.uninstall_method == "msi"

    def test_add_scrub_options(self) -> None:
        """add_scrub_options should add scrub flags."""
        parser = argparse.ArgumentParser()
        cli_help.add_scrub_options(parser)
        args = parser.parse_args(["--scrub-level", "nuclear"])
        assert args.scrub_level == "nuclear"

    def test_add_output_options(self) -> None:
        """add_output_options should add output flags."""
        parser = argparse.ArgumentParser()
        cli_help.add_output_options(parser)
        args = parser.parse_args(["--quiet", "--json"])
        assert args.quiet is True
        assert args.json is True
