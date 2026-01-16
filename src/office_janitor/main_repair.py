"""!
@file main_repair.py
@brief Office repair command handlers for Office Janitor.
@details Handles auto-repair, C2R repair, ODT repair, and OEM configuration
execution modes.
"""

from __future__ import annotations

import logging
import pathlib
from typing import TYPE_CHECKING

from . import auto_repair as auto_repair_module
from . import confirm, repair
from .main_progress import progress, progress_fail, progress_ok, progress_skip

if TYPE_CHECKING:
    import argparse

__all__ = [
    "handle_auto_repair_mode",
    "handle_repair_odt_mode",
    "handle_repair_c2r_mode",
    "handle_repair_mode",
    "handle_oem_config_mode",
]


def handle_auto_repair_mode(
    args: argparse.Namespace,
    human_log: logging.Logger,
    machine_log: logging.Logger,
) -> int:
    """!
    @brief Handle auto-repair mode - intelligent repair of all Office installations.
    @details Detects all Office installations and repairs them using the most
    appropriate method for each installation type.
    @param args Parsed command-line arguments.
    @param human_log Human-readable logger.
    @param machine_log Machine-readable (JSONL) logger.
    @returns Exit code (0 for success, non-zero for failure).
    """
    progress("-" * 60)
    progress("Auto-Repair Mode - Intelligent Office Repair")
    progress("-" * 60)

    dry_run = bool(getattr(args, "dry_run", False))
    culture = getattr(args, "repair_culture", None)
    platform = getattr(args, "repair_platform", None)
    silent = not getattr(args, "repair_visible", False)

    # Detect Office products
    progress("Detecting Office installations...", newline=False)
    products = auto_repair_module.detect_office_products()
    if not products:
        progress_fail("none found")
        human_log.warning("No Office installations detected for repair")
        print("\nNo Office installations detected.")
        print("If Office is installed, it may not be in a repairable state.")
        return 1
    progress_ok(f"{len(products)} found")

    # Display detected products
    progress("Detected products:")
    for product in products:
        progress(
            f"  • {product.product_name} ({product.version}, {product.install_type})",
            indent=1,
        )

    # Create repair plan
    progress("Creating repair plan...", newline=False)
    strategy = auto_repair_module.RepairStrategy.QUICK
    if getattr(args, "repair", None) == "full":
        strategy = auto_repair_module.RepairStrategy.FULL

    plan = auto_repair_module.create_repair_plan(
        products=products,
        strategy=strategy,
        dry_run=dry_run,
    )
    progress_ok()

    progress(f"  Method: {plan.recommended_method.value}", indent=1)
    progress(f"  Strategy: {plan.recommended_strategy.value}", indent=1)
    progress(f"  Estimated time: {plan.estimated_time_minutes} minutes", indent=1)

    # Show warnings
    for warning in plan.warnings:
        progress(f"  ⚠️  {warning}", indent=1)

    # Confirm with user unless forced
    if not dry_run and not getattr(args, "force", False):
        progress("Confirm repair?", newline=False)
        proceed = confirm.request_scrub_confirmation(dry_run=dry_run, force=False)
        if not proceed:
            progress_skip("cancelled by user")
            human_log.info("Auto-repair cancelled by user")
            return 0
        progress_ok()

    # Execute repair
    progress("=" * 60)
    progress("Executing auto-repair...")
    progress("=" * 60)

    result = auto_repair_module.execute_auto_repair(
        plan=plan,
        culture=culture,
        platform=platform,
        silent=silent,
        dry_run=dry_run,
    )

    progress("=" * 60)
    if result.success:
        progress(f"✓ {result.summary}")
        progress("=" * 60)
        print(f"\n{result.summary}")
        if result.products_repaired:
            print("\nRepaired products:")
            for prod in result.products_repaired:
                print(f"  ✓ {prod}")
        print("\nNote: A system restart may be required to complete the repair.")
        return 0
    else:
        progress(f"✗ {result.summary}")
        progress("=" * 60)
        print(f"\n{result.summary}")
        if result.products_repaired:
            print("\nRepaired products:")
            for prod in result.products_repaired:
                print(f"  ✓ {prod}")
        if result.products_failed:
            print("\nFailed products:")
            for prod in result.products_failed:
                print(f"  ✗ {prod}")
        if result.errors:
            print("\nErrors:")
            for error in result.errors:
                print(f"  - {error}")
        return 1


def handle_repair_odt_mode(
    args: argparse.Namespace,
    human_log: logging.Logger,
    machine_log: logging.Logger,
) -> int:
    """!
    @brief Handle repair via ODT configuration.
    @details Uses Office Deployment Tool setup.exe with a configuration file
    to repair/reconfigure Office installations.
    @param args Parsed command-line arguments.
    @param human_log Human-readable logger.
    @param machine_log Machine-readable (JSONL) logger.
    @returns Exit code (0 for success, non-zero for failure).
    """
    progress("-" * 60)
    progress("ODT Repair Mode")
    progress("-" * 60)

    dry_run = bool(getattr(args, "dry_run", False))
    preset = getattr(args, "repair_preset", None)
    config_path = getattr(args, "repair_config", None)

    # Determine config to use
    if config_path:
        progress(f"Using custom config: {config_path}")
        config_file = pathlib.Path(config_path)
        if not config_file.exists():
            progress(f"Configuration file not found: {config_path}", newline=False)
            progress_fail()
            return 1
    elif preset:
        progress(f"Using preset: {preset}")
        config_file = repair.get_oem_config_path(preset)
        if config_file is None:
            progress(f"Preset not found: {preset}", newline=False)
            progress_fail()
            progress("\nAvailable presets:")
            for name, _filename, exists in repair.list_oem_configs():
                if exists:
                    progress(f"  {name}", indent=1)
            return 1
    else:
        # Default to quick-repair preset
        preset = "quick-repair"
        progress(f"Using default preset: {preset}")
        config_file = repair.get_oem_config_path(preset)
        if config_file is None:
            progress("Default repair preset not found", newline=False)
            progress_fail()
            return 1

    # Check for setup.exe
    progress("Locating ODT setup.exe...", newline=False)
    setup_exe = repair.find_odt_setup_exe()
    if setup_exe is None:
        progress_fail("not found")
        human_log.error("ODT setup.exe not found")
        print("\nError: ODT setup.exe not found.")
        print("Download from: https://www.microsoft.com/en-us/download/details.aspx?id=49117")
        return 1
    progress_ok(str(setup_exe))

    # Confirm with user
    if not dry_run and not getattr(args, "force", False):
        progress("Confirm ODT repair?", newline=False)
        proceed = confirm.request_scrub_confirmation(dry_run=dry_run, force=False)
        if not proceed:
            progress_skip("cancelled by user")
            return 0
        progress_ok()

    # Execute repair
    progress("=" * 60)
    progress("Executing ODT repair...")
    progress("=" * 60)

    result = repair.reconfigure_office(
        config_file,
        dry_run=dry_run,
        timeout=getattr(args, "repair_timeout", 3600),
    )

    progress("=" * 60)
    if result.returncode == 0 or result.skipped:
        progress("ODT repair completed successfully")
        progress("=" * 60)
        if result.skipped:
            print(f"\n[DRY-RUN] Would execute: setup.exe /configure {config_file}")
        else:
            print("\n✓ ODT repair completed successfully.")
            print("\nNote: A system restart may be required.")
        return 0
    else:
        progress(f"ODT repair failed: {result.stderr or result.error}")
        progress("=" * 60)
        print("\n✗ ODT repair failed.")
        if result.stderr:
            print(f"\nError: {result.stderr}")
        return 1


def handle_repair_c2r_mode(
    args: argparse.Namespace,
    human_log: logging.Logger,
    machine_log: logging.Logger,
) -> int:
    """!
    @brief Handle repair via OfficeClickToRun.exe directly.
    @details Uses OfficeClickToRun.exe for granular control over C2R repair.
    @param args Parsed command-line arguments.
    @param human_log Human-readable logger.
    @param machine_log Machine-readable (JSONL) logger.
    @returns Exit code (0 for success, non-zero for failure).
    """
    progress("-" * 60)
    progress("C2R Direct Repair Mode")
    progress("-" * 60)

    dry_run = bool(getattr(args, "dry_run", False))
    culture = getattr(args, "repair_culture", "en-us")
    platform = getattr(args, "repair_platform", None)
    silent = not getattr(args, "repair_visible", False)

    # Check if C2R Office is installed
    progress("Checking for C2R installation...", newline=False)
    if not repair.is_c2r_office_installed():
        progress_fail("not found")
        human_log.error("No Click-to-Run Office installation detected")
        print("\nError: No Click-to-Run Office installation found.")
        print("This repair mode only works with Office C2R installations.")
        return 1
    progress_ok()

    # Get C2R info
    progress("Gathering installation details...", newline=False)
    c2r_info = repair.get_installed_c2r_info()
    progress_ok()
    progress(f"  Version: {c2r_info.get('version', 'unknown')}", indent=1)
    progress(f"  Platform: {c2r_info.get('platform', 'unknown')}", indent=1)
    progress(f"  Culture: {c2r_info.get('culture', 'unknown')}", indent=1)

    # Locate OfficeClickToRun.exe
    progress("Locating OfficeClickToRun.exe...", newline=False)
    exe_path = repair.find_officeclicktorun_exe()
    if exe_path is None:
        progress_fail("not found")
        human_log.error("OfficeClickToRun.exe not found")
        return 1
    progress_ok(str(exe_path))

    # Determine repair type
    repair_type = getattr(args, "repair", "quick") or "quick"
    progress(f"Repair type: {repair_type.upper()}")

    if repair_type == "full":
        progress("\n⚠️  WARNING: Full repair may reinstall excluded applications!")
        config = repair.RepairConfig.full_repair(
            platform=platform,
            culture=culture,
            silent=silent,
        )
    else:
        config = repair.RepairConfig.quick_repair(
            platform=platform,
            culture=culture,
            silent=silent,
        )

    # Confirm with user
    if not dry_run and not getattr(args, "force", False):
        progress("Confirm C2R repair?", newline=False)
        proceed = confirm.request_scrub_confirmation(dry_run=dry_run, force=False)
        if not proceed:
            progress_skip("cancelled by user")
            return 0
        progress_ok()

    # Execute repair
    progress("=" * 60)
    progress(f"Executing {repair_type.upper()} C2R repair...")
    progress("=" * 60)

    result = repair.run_repair(config, dry_run=dry_run)

    progress("=" * 60)
    if result.success or result.skipped:
        progress(f"✓ {result.summary}")
        progress("=" * 60)
        if result.skipped:
            print(f"\n[DRY-RUN] {result.summary}")
        else:
            print(f"\n✓ {result.summary}")
            print("\nNote: A system restart may be required.")
        return 0
    else:
        progress(f"✗ {result.summary}")
        progress("=" * 60)
        print(f"\n✗ {result.summary}")
        if result.stderr:
            print(f"\nError: {result.stderr}")
        return 1


def handle_repair_mode(
    args: argparse.Namespace,
    mode: str,
    human_log: logging.Logger,
    machine_log: logging.Logger,
) -> int:
    """!
    @brief Handle Office repair operations.
    @details Dispatches to quick/full repair or custom XML configuration based
    on the mode string and command-line arguments.
    @param args Parsed command-line arguments.
    @param mode Mode string (repair:quick, repair:full, or repair:config).
    @param human_log Human-readable logger.
    @param machine_log Machine-readable (JSONL) logger.
    @returns Exit code (0 for success, non-zero for failure).
    """
    progress("-" * 60)
    progress("Office Click-to-Run Repair Mode")
    progress("-" * 60)

    dry_run = bool(getattr(args, "dry_run", False))
    culture = getattr(args, "repair_culture", "en-us")
    platform = getattr(args, "repair_platform", None)
    silent = not getattr(args, "repair_visible", False)

    # Check if C2R Office is installed
    progress("Checking for Click-to-Run installation...", newline=False)
    if not repair.is_c2r_office_installed():
        progress_fail("not found")
        human_log.error("No Click-to-Run Office installation detected")
        machine_log.info(
            "repair.error",
            extra={"event": "repair.error", "error": "c2r_not_installed"},
        )
        print("\nError: No Click-to-Run Office installation found.")
        print("This repair option only works with Office C2R installations.")
        print("\nFor MSI-based installations, use the standard Windows repair:")
        print("  Control Panel > Programs > Programs and Features > [Office] > Change > Repair")
        return 1
    progress_ok()

    # Get installed Office info
    progress("Gathering installation details...", newline=False)
    c2r_info = repair.get_installed_c2r_info()
    progress_ok()
    progress(f"  Version: {c2r_info.get('version', 'unknown')}", indent=1)
    progress(f"  Platform: {c2r_info.get('platform', 'unknown')}", indent=1)
    progress(f"  Culture: {c2r_info.get('culture', 'unknown')}", indent=1)
    progress(f"  Products: {c2r_info.get('product_ids', 'unknown')}", indent=1)

    # Handle custom XML configuration
    if mode == "repair:config":
        config_path = pathlib.Path(getattr(args, "repair_config", ""))
        if not config_path.exists():
            progress(f"Configuration file not found: {config_path}", newline=False)
            progress_fail()
            return 1
        progress(f"Using custom configuration: {config_path}")
        result = repair.reconfigure_office(
            config_path,
            dry_run=dry_run,
        )
        if result.returncode == 0 or result.skipped:
            progress("Reconfiguration completed successfully", newline=False)
            progress_ok()
            return 0
        progress(f"Reconfiguration failed: {result.stderr or result.error}", newline=False)
        progress_fail()
        return 1

    # Determine repair type
    repair_type_str = mode.split(":")[-1]  # quick or full
    progress(f"Repair type: {repair_type_str.upper()}")

    if repair_type_str == "full":
        progress("\n⚠️  WARNING: Full Online Repair may reinstall excluded applications!")
        progress("    This operation requires internet connectivity and may take 30-60 minutes.\n")
        config = repair.RepairConfig.full_repair(
            platform=platform,
            culture=culture,
            silent=silent,
        )
    else:
        progress("Quick Repair runs locally and typically completes in 5-15 minutes.")
        config = repair.RepairConfig.quick_repair(
            platform=platform,
            culture=culture,
            silent=silent,
        )

    # Confirm with user unless in auto mode
    if not dry_run and not getattr(args, "force", False):
        progress("Confirm repair operation?", newline=False)
        proceed = confirm.request_scrub_confirmation(dry_run=dry_run, force=False)
        if not proceed:
            progress_skip("cancelled by user")
            human_log.info("Repair cancelled by user")
            return 0
        progress_ok()

    # Execute repair
    progress("=" * 60)
    progress(f"Executing {repair_type_str.upper()} repair...")
    progress("=" * 60)

    repair_result = repair.run_repair(config, dry_run=dry_run)

    progress("=" * 60)
    if repair_result.success or repair_result.skipped:
        progress(f"Repair completed: {repair_result.summary}")
        progress("=" * 60)
        if repair_result.skipped:
            print(f"\n[DRY-RUN] {repair_result.summary}")
        else:
            print(f"\n✓ {repair_result.summary}")
            print("\nNote: A system restart may be required to complete the repair.")
        return 0
    else:
        progress(f"Repair failed: {repair_result.summary}")
        progress("=" * 60)
        print(f"\n✗ {repair_result.summary}")
        if repair_result.stderr:
            print(f"\nError details:\n{repair_result.stderr}")
        return 1


def handle_oem_config_mode(
    args: argparse.Namespace,
    mode: str,
    human_log: logging.Logger,
    machine_log: logging.Logger,
) -> int:
    """!
    @brief Handle OEM configuration execution.
    @details Executes a bundled or custom XML configuration using ODT setup.exe.
    @param args Parsed command-line arguments.
    @param mode Mode string (oem-config:<preset-name>).
    @param human_log Human-readable logger.
    @param machine_log Machine-readable (JSONL) logger.
    @returns Exit code (0 for success, non-zero for failure).
    """
    progress("-" * 60)
    progress("OEM Configuration Mode")
    progress("-" * 60)

    dry_run = bool(getattr(args, "dry_run", False))
    preset_name = mode.split(":", 1)[-1] if ":" in mode else getattr(args, "oem_config", "")

    # List available presets if none specified
    if not preset_name:
        progress("Available OEM configuration presets:")
        for name, filename, exists in repair.list_oem_configs():
            status = "✓" if exists else "✗ (missing)"
            progress(f"  {name}: {filename} {status}", indent=1)
        return 0

    # Resolve the config path
    config_path = repair.get_oem_config_path(preset_name)
    if config_path is None:
        progress(f"OEM config not found: {preset_name}", newline=False)
        progress_fail()
        human_log.error(f"OEM config preset not found: {preset_name}")
        machine_log.info(
            "oem_config.error",
            extra={"event": "oem_config.error", "preset": preset_name, "error": "not_found"},
        )
        progress("\nAvailable presets:")
        for name, _filename, exists in repair.list_oem_configs():
            if exists:
                progress(f"  {name}", indent=1)
        return 1

    progress(f"Preset: {preset_name}")
    progress(f"Config file: {config_path}")

    # Check for ODT setup.exe
    setup_exe = repair.find_odt_setup_exe()
    if setup_exe is None:
        progress("ODT setup.exe not found", newline=False)
        progress_fail()
        human_log.error("ODT setup.exe not found")
        machine_log.info(
            "oem_config.error",
            extra={"event": "oem_config.error", "error": "setup_not_found"},
        )
        print("\nError: ODT setup.exe not found.")
        print("Please ensure setup.exe is in the oem/ folder or download it from:")
        print("  https://www.microsoft.com/en-us/download/details.aspx?id=49117")
        return 1

    progress(f"Setup.exe: {setup_exe}")

    # Warn about destructive operations
    if preset_name in ("full-removal",):
        progress("\n⚠️  WARNING: This will REMOVE all Office installations!")
        progress("    This action cannot be undone.\n")
    elif "repair" in preset_name.lower():
        progress("\nNote: Repair operations may take 5-60 minutes depending on type.\n")

    # Confirm with user unless forced
    if not dry_run and not getattr(args, "force", False):
        progress("Confirm operation?", newline=False)
        proceed = confirm.request_scrub_confirmation(dry_run=dry_run, force=False)
        if not proceed:
            progress_skip("cancelled by user")
            human_log.info("OEM config cancelled by user")
            return 0
        progress_ok()

    # Execute
    progress("=" * 60)
    progress(f"Executing configuration: {preset_name}")
    progress("=" * 60)

    result = repair.run_oem_config(
        preset_name,
        dry_run=dry_run,
    )

    progress("=" * 60)
    if result.returncode == 0 or result.skipped:
        progress("Configuration completed successfully")
        progress("=" * 60)
        if result.skipped:
            print(f"\n[DRY-RUN] Would execute: setup.exe /configure {config_path}")
        else:
            print(f"\n✓ Configuration '{preset_name}' applied successfully.")
            print("\nNote: A system restart may be required to complete changes.")
        return 0
    else:
        progress(f"Configuration failed: {result.stderr or result.error}")
        progress("=" * 60)
        print(f"\n✗ Configuration '{preset_name}' failed.")
        if result.stderr:
            print(f"\nError details:\n{result.stderr}")
        return 1
