"""!
@file main_odt.py
@brief Office Deployment Tool (ODT) command handlers for Office Janitor.
@details Handles ODT listing commands, configuration building, installation,
and download operations.
"""

from __future__ import annotations

import pathlib
from typing import TYPE_CHECKING

from . import odt_build
from .main_progress import progress, progress_fail, progress_ok

if TYPE_CHECKING:
    import argparse

__all__ = [
    "handle_odt_list_commands",
    "handle_odt_build_commands",
]


def handle_odt_list_commands(args: argparse.Namespace) -> bool:
    """!
    @brief Handle ODT listing commands that don't require elevation.
    @details Processes --odt-list-* commands which are read-only operations
    that can run without administrative privileges.
    @param args Parsed command-line arguments.
    @returns True if a list command was handled (caller should exit), False otherwise.
    """
    if getattr(args, "odt_list_products", False):
        print("\nAvailable Office Products for ODT Configuration:")
        print("=" * 80)
        for product in odt_build.list_products():
            channels_list: list[str] = product.get("channels", [])
            channels = ", ".join(channels_list)
            print(f"\n  {product['id']}")
            print(f"      Name: {product['name']}")
            print(f"      Channels: {channels}")
            print(f"      {product['description']}")
        return True

    if getattr(args, "odt_list_presets", False):
        print("\nAvailable Installation Presets:")
        print("=" * 80)
        for preset in odt_build.list_presets():
            products_list: list[str] = preset.get("products", [])
            products = ", ".join(products_list)
            print(f"\n  {preset['name']}")
            print(f"      Products: {products}")
            print(f"      Architecture: {preset['architecture']}, Channel: {preset['channel']}")
            print(f"      {preset['description']}")
        return True

    if getattr(args, "odt_list_channels", False):
        print("\nAvailable Update Channels:")
        print("=" * 60)
        for ch in odt_build.list_channels():
            print(f"  {ch['name']:<30} {ch['value']}")
        return True

    if getattr(args, "odt_list_languages", False):
        print("\nSupported Language Codes:")
        print("=" * 60)
        langs = odt_build.list_languages()
        for i in range(0, len(langs), 4):
            row = langs[i : i + 4]
            print("  " + "  ".join(f"{lang:<14}" for lang in row))
        return True

    return False


def handle_odt_build_commands(args: argparse.Namespace) -> bool:
    """!
    @brief Handle ODT build and configuration generation commands.
    @details Processes --odt-build, --odt-install, --odt-removal, --odt-download,
    and author aliases (--goobler, --pupa).
    These commands may write files and run after elevation.
    @param args Parsed command-line arguments.
    @returns True if an ODT command was handled (caller should exit), False otherwise.
    """
    # Handle author quick install aliases
    if getattr(args, "goobler", False):
        return _run_author_install(
            args,
            preset="ltsc2024-full-x64-clean",
            languages=["pt-pt", "en-us"],
            name="Goobler",
        )

    if getattr(args, "pupa", False):
        return _run_author_install(
            args,
            preset="ltsc2024-x64-clean",
            languages=["pt-pt", "en-us"],
            name="Pupa",
        )

    # Handle install command (actually runs ODT)
    if getattr(args, "odt_install", False):
        return _run_odt_install(args)

    # Handle build command (generates XML)
    if getattr(args, "odt_build", False):
        return _build_odt_config(args)

    # Handle removal XML generation
    if getattr(args, "odt_removal", False):
        return _build_odt_removal(args)

    # Handle download XML generation
    if getattr(args, "odt_download", None):
        return _build_odt_download(args)

    return False


def _run_odt_install(args: argparse.Namespace) -> bool:
    """!
    @brief Run ODT to install Office with the specified configuration.
    @param args Parsed command-line arguments.
    @returns True (command was handled).
    """
    dry_run = getattr(args, "dry_run", False)

    try:
        preset = getattr(args, "odt_preset", None)
        products = getattr(args, "odt_products", None)
        languages = getattr(args, "odt_languages", None) or ["en-us"]

        if preset:
            config = odt_build.ODTConfig.from_preset(preset, languages)
        elif products:
            arch = (
                odt_build.Architecture.X64
                if getattr(args, "odt_arch", "64") == "64"
                else odt_build.Architecture.X86
            )

            channel_arg = getattr(args, "odt_channel", None)
            channel = odt_build.UpdateChannel.CURRENT
            if channel_arg:
                for ch in odt_build.UpdateChannel:
                    if ch.name.lower() == channel_arg.lower().replace("-", "_"):
                        channel = ch
                        break
                    if ch.value.lower() == channel_arg.lower():
                        channel = ch
                        break

            exclude_apps = getattr(args, "odt_exclude_apps", None) or []
            product_configs = [
                odt_build.ProductConfig(pid, languages=languages, exclude_apps=exclude_apps)
                for pid in products
            ]

            if getattr(args, "odt_include_visio", False):
                product_configs.append(
                    odt_build.ProductConfig("VisioProRetail", languages=languages)
                )
            if getattr(args, "odt_include_project", False):
                product_configs.append(
                    odt_build.ProductConfig("ProjectProRetail", languages=languages)
                )

            config = odt_build.ODTConfig(
                products=product_configs,
                architecture=arch,
                channel=channel,
                shared_computer_licensing=getattr(args, "odt_shared_computer", False),
                remove_msi=getattr(args, "odt_remove_msi", False),
            )
        else:
            progress("No preset or products specified for installation.")
            progress("  Use --odt-preset or --odt-product to specify what to install.", indent=1)
            progress("  Use --odt-list-presets or --odt-list-products for options.", indent=1)
            return True

        # Show what we're about to install using consistent logging format
        progress("-" * 60)
        progress("ODT Installation")
        progress("-" * 60)
        progress(f"  Products: {', '.join(p.product_id for p in config.products)}")
        progress(
            f"  Languages: {', '.join(config.products[0].languages) if config.products else 'none'}"
        )
        progress(f"  Architecture: {config.architecture.value}-bit")
        progress(f"  Channel: {config.channel.value}")
        progress("-" * 60)

        # Run installation - spinner handles progress display, don't use incomplete line
        result = odt_build.run_odt_install(config, dry_run=dry_run)

        if result.success:
            progress_ok(f"{result.duration:.1f}s")
            progress("Installation complete")
        else:
            progress_fail(f"exit {result.return_code}")
            if result.stderr:
                progress(f"  Error: {result.stderr}", indent=1)
            if result.config_path:
                progress(f"  Config preserved: {result.config_path}", indent=1)

        return True

    except ValueError as e:
        progress_fail(str(e))
        return True
    except FileNotFoundError as e:
        progress_fail(str(e))
        return True


def _run_author_install(
    args: argparse.Namespace,
    *,
    preset: str,
    languages: list[str],
    name: str,
) -> bool:
    """!
    @brief Run a pre-configured author install alias.
    @param args Parsed command-line arguments.
    @param preset The preset name to use.
    @param languages List of language codes.
    @param name Display name for the alias.
    @returns True (command was handled).
    """
    dry_run = getattr(args, "dry_run", False)

    try:
        config = odt_build.ODTConfig.from_preset(preset, languages)

        # Show what we're about to install using consistent logging format
        progress("-" * 60)
        progress(f"Quick Install: {name}")
        progress("-" * 60)
        progress(f"  Preset: {preset}")
        progress(f"  Products: {', '.join(p.product_id for p in config.products)}")
        progress(f"  Languages: {', '.join(languages)}")
        progress(f"  Architecture: {config.architecture.value}-bit")
        progress(f"  Channel: {config.channel.value}")
        if config.products and config.products[0].exclude_apps:
            progress(f"  Excluded: {', '.join(config.products[0].exclude_apps)}")
        progress("-" * 60)

        # Run installation - spinner handles progress display, don't use incomplete line
        result = odt_build.run_odt_install(config, dry_run=dry_run)

        if result.success:
            progress_ok(f"{result.duration:.1f}s")
            progress(f"Installation complete: {name}")
        else:
            progress_fail(f"exit {result.return_code}")
            if result.stderr:
                progress(f"  Error: {result.stderr}", indent=1)
            if result.config_path:
                progress(f"  Config preserved: {result.config_path}", indent=1)

        return True

    except ValueError as e:
        progress_fail(str(e))
        return True
    except FileNotFoundError as e:
        progress_fail(str(e))
        return True


def _build_odt_config(args: argparse.Namespace) -> bool:
    """!
    @brief Build ODT installation XML configuration.
    @param args Parsed command-line arguments.
    @returns True (command was handled).
    """
    try:
        preset = getattr(args, "odt_preset", None)
        products = getattr(args, "odt_products", None)
        languages = getattr(args, "odt_languages", None) or ["en-us"]

        if preset:
            # Use preset configuration
            config = odt_build.ODTConfig.from_preset(preset, languages)
        elif products:
            # Build custom configuration from product list
            arch = (
                odt_build.Architecture.X64
                if getattr(args, "odt_arch", "64") == "64"
                else odt_build.Architecture.X86
            )

            # Determine channel
            channel_arg = getattr(args, "odt_channel", None)
            channel = odt_build.UpdateChannel.CURRENT
            if channel_arg:
                # Try to match channel by name or value
                for ch in odt_build.UpdateChannel:
                    if ch.name.lower() == channel_arg.lower().replace("-", "_"):
                        channel = ch
                        break
                    if ch.value.lower() == channel_arg.lower():
                        channel = ch
                        break

            exclude_apps = getattr(args, "odt_exclude_apps", None) or []
            product_configs = [
                odt_build.ProductConfig(pid, languages=languages, exclude_apps=exclude_apps)
                for pid in products
            ]

            # Add optional products
            if getattr(args, "odt_include_visio", False):
                product_configs.append(
                    odt_build.ProductConfig("VisioProRetail", languages=languages)
                )
            if getattr(args, "odt_include_project", False):
                product_configs.append(
                    odt_build.ProductConfig("ProjectProRetail", languages=languages)
                )

            config = odt_build.ODTConfig(
                products=product_configs,
                architecture=arch,
                channel=channel,
                shared_computer_licensing=getattr(args, "odt_shared_computer", False),
                remove_msi=getattr(args, "odt_remove_msi", False),
            )
        else:
            # No preset or products specified - use default M365 ProPlus
            print("No preset or products specified. Use --odt-preset or --odt-product.")
            print("Use --odt-list-presets or --odt-list-products to see available options.")
            return True

        xml_output = odt_build.build_xml(config)

        output_path = getattr(args, "odt_output", None)
        if output_path:
            pathlib.Path(output_path).write_text(xml_output, encoding="utf-8")
            print(f"ODT configuration written to: {output_path}")
        else:
            print(xml_output)

    except ValueError as e:
        print(f"Error: {e}", file=__import__("sys").stderr)
        return True

    return True


def _build_odt_removal(args: argparse.Namespace) -> bool:
    """!
    @brief Build ODT removal XML configuration.
    @param args Parsed command-line arguments.
    @returns True (command was handled).
    """
    product_ids = getattr(args, "odt_products", None)
    force_shutdown = not getattr(args, "no_force_app_shutdown", False)
    remove_msi = getattr(args, "odt_remove_msi", False)

    xml_output = odt_build.build_removal_xml(
        remove_all=not product_ids,
        product_ids=product_ids,
        force_app_shutdown=force_shutdown,
        remove_msi=remove_msi,
    )

    output_path = getattr(args, "odt_output", None)
    if output_path:
        pathlib.Path(output_path).write_text(xml_output, encoding="utf-8")
        print(f"ODT removal configuration written to: {output_path}")
    else:
        print(xml_output)

    return True


def _build_odt_download(args: argparse.Namespace) -> bool:
    """!
    @brief Build ODT download XML configuration.
    @param args Parsed command-line arguments.
    @returns True (command was handled).
    """
    download_path = getattr(args, "odt_download", None)
    if not download_path:
        return False

    try:
        preset = getattr(args, "odt_preset", None)
        products = getattr(args, "odt_products", None)
        languages = getattr(args, "odt_languages", None) or ["en-us"]

        if preset:
            config = odt_build.ODTConfig.from_preset(preset, languages)
        elif products:
            arch = (
                odt_build.Architecture.X64
                if getattr(args, "odt_arch", "64") == "64"
                else odt_build.Architecture.X86
            )
            channel = odt_build.UpdateChannel.CURRENT
            channel_arg = getattr(args, "odt_channel", None)
            if channel_arg:
                for ch in odt_build.UpdateChannel:
                    if ch.name.lower() == channel_arg.lower().replace("-", "_"):
                        channel = ch
                        break
                    if ch.value.lower() == channel_arg.lower():
                        channel = ch
                        break

            product_configs = [
                odt_build.ProductConfig(pid, languages=languages) for pid in products
            ]
            config = odt_build.ODTConfig(
                products=product_configs,
                architecture=arch,
                channel=channel,
            )
        else:
            # Default to M365 ProPlus
            config = odt_build.ODTConfig.from_preset("365-proplus-x64", languages)

        xml_output = odt_build.build_download_xml(config, download_path)

        output_path = getattr(args, "odt_output", None)
        if output_path:
            pathlib.Path(output_path).write_text(xml_output, encoding="utf-8")
            print(f"ODT download configuration written to: {output_path}")
        else:
            print(xml_output)

    except ValueError as e:
        print(f"Error: {e}", file=__import__("sys").stderr)

    return True
