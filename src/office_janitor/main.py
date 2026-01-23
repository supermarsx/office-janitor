"""!
@file main.py
@brief Primary CLI entry point for the Office Janitor application.
@details Implements the specification surface area by wiring argument parsing
into detection, planning, and scrub routines while managing bootstrap tasks
like elevation, logging, and console configuration.
"""

from __future__ import annotations

import argparse
import logging
import os
import pathlib
import signal
import sys
import time
from collections.abc import Iterable

from . import (
    cli_help,
    confirm,
    detect,  # noqa: F401 - re-exported for test patching
    elevation,
    exec_utils,
    logging_ext,
    safety,
    scrub,
    spinner,
    tui,
    ui,
    version,
)
from . import plan as plan_module
from .app_state import AppState  # noqa: F401

# Import from refactored modules
from .main_odt import handle_odt_build_commands, handle_odt_list_commands
from .main_progress import (
    enable_vt_mode_if_possible,
    get_elapsed_secs,
    progress,
    progress_fail,
    progress_ok,
    progress_skip,
    set_main_start_time,
)
from .main_repair import (
    handle_auto_repair_mode,
    handle_oem_config_mode,
    handle_repair_c2r_mode,
    handle_repair_mode,
    handle_repair_odt_mode,
)
from .main_state import (
    build_app_state,
    collect_plan_options,
    determine_mode,
    enforce_runtime_guards,
    handle_plan_artifacts,
    load_config_file,
    resolve_log_directory,
    run_detection,
    should_use_tui,
)

# Re-export with underscore prefix for backwards compatibility (tests patch these)
_progress = progress
_progress_ok = progress_ok
_progress_fail = progress_fail
_progress_skip = progress_skip
_get_elapsed_secs = get_elapsed_secs
_determine_mode = determine_mode
_should_use_tui = should_use_tui
_build_app_state = build_app_state
_collect_plan_options = collect_plan_options
_run_detection = run_detection
_handle_plan_artifacts = handle_plan_artifacts
_enforce_runtime_guards = enforce_runtime_guards
_resolve_log_directory = resolve_log_directory
_load_config_file = load_config_file
_handle_odt_list_commands = handle_odt_list_commands
_handle_odt_build_commands = handle_odt_build_commands
_handle_auto_repair_mode = handle_auto_repair_mode
_handle_repair_odt_mode = handle_repair_odt_mode
_handle_repair_c2r_mode = handle_repair_c2r_mode
_handle_repair_mode = handle_repair_mode
_handle_oem_config_mode = handle_oem_config_mode

# Module-level reference to start time for signal handlers
_MAIN_START_TIME: float = time.perf_counter()


def ensure_admin_and_relaunch_if_needed() -> None:
    """!
    @brief Request elevation if the current process lacks administrative rights.
    """
    if elevation.is_admin():
        return

    argv = list(sys.argv)
    try:
        executable_path = pathlib.Path(sys.executable).resolve()
        argv0_path = pathlib.Path(argv[0]).resolve() if argv else None
    except Exception:
        executable_path = None
        argv0_path = None

    if executable_path is not None and argv0_path == executable_path:
        argv = argv[1:]

    if not elevation.relaunch_as_admin(argv):
        raise SystemExit("Failed to request elevation via ShellExecuteW.")
    sys.exit(0)


def build_arg_parser() -> argparse.ArgumentParser:
    """!
    @brief Create the top-level argument parser with the specification's surface area.
    @details The parser wires in the mutually exclusive modes and shared options
    defined in the specification so later feature work can hook into the parsed
    values without changing the public CLI signature.
    """
    metadata = version.build_info()
    return cli_help.build_arg_parser(version_info=metadata)


def _bootstrap_logging(
    args: argparse.Namespace,
) -> tuple[logging.Logger, logging.Logger]:
    """!
    @brief Initialize human and machine loggers using :mod:`logging_ext` helpers.
    @returns A tuple of configured human and machine loggers.
    """
    logdir = _resolve_log_directory(getattr(args, "logdir", None))
    logdir.mkdir(parents=True, exist_ok=True)
    args.logdir = str(logdir)
    human_logger, machine_logger = logging_ext.setup_logging(
        logdir,
        json_to_stdout=getattr(args, "json", False),
    )
    args.human_logger = human_logger
    args.machine_logger = machine_logger
    if getattr(args, "quiet", False):
        human_logger.setLevel(logging.ERROR)
    return human_logger, machine_logger


def _handle_shutdown_signal(signum: int, frame: object) -> None:
    """!
    @brief Handle Ctrl+C (SIGINT) for clean shutdown.
    """
    sig_name = signal.Signals(signum).name if hasattr(signal, "Signals") else str(signum)
    elapsed = get_elapsed_secs()
    print(flush=True)  # Newline after any partial output
    print(f"[{elapsed:12.6f}] " + "=" * 50, flush=True)
    print(
        f"[{elapsed:12.6f}] \033[33mShutting down...\033[0m (received {sig_name})",
        flush=True,
    )
    print(f"[{elapsed:12.6f}] Cleanup in progress, please wait...", flush=True)
    print(f"[{elapsed:12.6f}] " + "=" * 50, flush=True)
    sys.exit(130)  # Standard exit code for SIGINT


def _print_shutdown_message() -> None:
    """Print a clean shutdown message with timestamp."""
    elapsed = get_elapsed_secs()
    print(flush=True)  # Newline after any partial output
    print(f"[{elapsed:12.6f}] " + "=" * 50, flush=True)
    print(
        f"[{elapsed:12.6f}] \033[33mShutting down...\033[0m (keyboard interrupt)",
        flush=True,
    )
    print(f"[{elapsed:12.6f}] Cleanup in progress, please wait...", flush=True)
    print(f"[{elapsed:12.6f}] " + "=" * 50, flush=True)


def main(argv: Iterable[str] | None = None, *, start_time: float | None = None) -> int:
    """!
    @brief Entry point invoked by the shim and PyInstaller bundle.
    @param start_time Optional startup timestamp from entry point for continuous timing.
    @returns Process exit code integer.
    """
    global _MAIN_START_TIME
    # Use provided start_time for continuous timestamps, or start fresh
    _MAIN_START_TIME = start_time if start_time is not None else time.perf_counter()
    set_main_start_time(_MAIN_START_TIME)

    # Check for elevation marker and set environment flag
    argv_list = list(argv) if argv is not None else None
    if argv_list is None:
        argv_list = sys.argv[1:]
    if argv_list and argv_list[0] == "--_elevated-marker":
        os.environ["OFFICE_JANITOR_ELEVATED"] = "1"
        argv_list = argv_list[1:]  # Remove the marker from args

    # Install signal handler for clean Ctrl+C shutdown
    signal.signal(signal.SIGINT, _handle_shutdown_signal)
    if hasattr(signal, "SIGBREAK"):  # Windows-specific
        signal.signal(signal.SIGBREAK, _handle_shutdown_signal)

    try:
        exit_code = _main_impl(argv_list)
    except KeyboardInterrupt:
        _print_shutdown_message()
        exit_code = 130  # Standard exit code for SIGINT

    # Keep console open if this was an auto-elevated process
    elevation.pause_if_elevated(exit_code)
    return exit_code


def _main_impl(argv: Iterable[str] | None = None) -> int:
    """!
    @brief Implementation of main() wrapped with KeyboardInterrupt handling.
    """
    progress("=" * 60)
    progress("Office Janitor - Main Entry Point")
    progress("=" * 60)

    # Phase 0: Pre-parse for ODT listing commands (no elevation needed)
    parser = build_arg_parser()
    args = parser.parse_args(list(argv) if argv is not None else None)

    # Handle ODT listing commands early (before elevation check)
    if handle_odt_list_commands(args):
        return 0

    # Phase 1: Elevation check
    progress("Phase 1: Checking administrative privileges...", newline=False)
    try:
        ensure_admin_and_relaunch_if_needed()
        progress_ok("elevated")
    except SystemExit:
        progress_fail("elevation required")
        raise

    # Phase 2: Console setup
    progress("Phase 2: Configuring console...", newline=False)
    enable_vt_mode_if_possible()
    progress_ok("VT mode enabled")

    # Handle ODT build commands (after elevation since they write files)
    if handle_odt_build_commands(args):
        return 0

    # Phase 3: Timeout configuration
    progress("Phase 3: Configuring execution timeout...", newline=False)
    timeout_val = getattr(args, "timeout", None)
    exec_utils.set_global_timeout(timeout_val)
    progress_ok(f"{timeout_val}s" if timeout_val else "default")

    # Phase 5: Logging bootstrap
    progress("Phase 5: Initializing logging subsystem...", newline=False)
    human_log, machine_log = _bootstrap_logging(args)
    if getattr(args, "quiet", False):
        human_log.setLevel(logging.ERROR)
    progress_ok(f"logdir={getattr(args, 'logdir', 'default')}")

    # Phase 6: Mode determination
    progress("Phase 6: Determining operation mode...", newline=False)
    mode = determine_mode(args)
    progress_ok(mode)

    # Log startup event
    progress("Emitting startup telemetry...")
    machine_log.info(
        "startup",
        extra={
            "event": "startup",
            "data": {"mode": mode, "dry_run": bool(getattr(args, "dry_run", False))},
        },
    )

    # Phase 7: App state construction
    progress("Phase 7: Building application state...", newline=False)
    app_state = build_app_state(args, human_log, machine_log, start_time=_MAIN_START_TIME)
    progress_ok()

    # ---------------------------------------------------------------------------
    # Install mode handling (new subcommand)
    # ---------------------------------------------------------------------------
    if mode.startswith("install:"):
        progress("Entering install mode...")
        # Route install subcommand modes to ODT handlers
        # The determine_mode already mapped subcommand options back to ODT args
        if handle_odt_build_commands(args):
            return 0
        progress("Install mode requires --preset or --product specification.")
        return 1

    # ---------------------------------------------------------------------------
    # Repair mode handling
    # ---------------------------------------------------------------------------
    # Auto-repair mode handling - intelligent repair of all Office installations
    if mode == "auto-repair":
        progress("Entering auto-repair mode...")
        return handle_auto_repair_mode(args, human_log, machine_log)

    # Repair-ODT mode handling - repair via ODT configuration
    if mode == "repair-odt":
        progress("Entering ODT repair mode...")
        return handle_repair_odt_mode(args, human_log, machine_log)

    # Repair-C2R mode handling - repair via OfficeClickToRun.exe
    if mode == "repair-c2r":
        progress("Entering C2R repair mode...")
        return handle_repair_c2r_mode(args, human_log, machine_log)

    # Repair mode handling - separate from standard detection/scrub flow
    if mode.startswith("repair:"):
        progress("Entering repair mode...")
        return handle_repair_mode(args, mode, human_log, machine_log)

    # ---------------------------------------------------------------------------
    # Remove mode handling (maps to auto-all, target, etc.)
    # ---------------------------------------------------------------------------
    if mode.startswith("remove:"):
        # remove:msi-only and remove:c2r-only map to target-specific removal
        progress(f"Entering remove mode ({mode.split(':')[1]})...")
        # Fall through to standard scrub flow - mode already set correctly

    # OEM config mode handling - execute bundled XML configurations
    if mode.startswith("oem-config:"):
        progress("Entering OEM configuration mode...")
        return handle_oem_config_mode(args, mode, human_log, machine_log)

    # Interactive mode handling
    if mode == "interactive":
        progress("Entering interactive mode...")
        if getattr(args, "tui", False):
            progress("Launching TUI (forced via --tui)...")
            tui.run_tui(app_state)
        else:
            tui_candidate = should_use_tui(args)
            if tui_candidate:
                progress("Launching TUI (auto-detected)...")
                tui.run_tui(app_state)
            else:
                progress("Launching CLI interface...")
                ui.run_cli(app_state)
        progress("Interactive session complete.")
        return 0

    progress("-" * 60)
    progress(f"Running in non-interactive mode: {mode}")
    progress("-" * 60)

    # Start spinner for non-interactive mode
    spinner.start_spinner_thread()

    # Phase 8: Detection
    spinner.set_task("Detecting Office installations")
    progress("Phase 8: Running Office detection...")
    logdir_path = pathlib.Path(getattr(args, "logdir", _resolve_log_directory(None))).expanduser()
    limited_flag = bool(getattr(args, "limited_user", False))
    if limited_flag:
        progress("Using limited user token for detection", indent=1)
    inventory = run_detection(
        machine_log,
        logdir_path,
        limited_user=limited_flag or None,
    )
    item_count = sum(len(v) if hasattr(v, "__len__") else 0 for v in inventory.values())
    progress(f"Detection complete: {item_count} items found", indent=1)

    # Phase 9: Plan generation
    spinner.set_task("Building execution plan")
    progress("Phase 9: Building execution plan...")
    options = collect_plan_options(args, mode)
    progress(
        f"Plan options: dry_run={options.get('dry_run')}, force={options.get('force')}",
        indent=1,
    )
    generated_plan = plan_module.build_plan(inventory, options)
    progress(f"Generated {len(generated_plan)} plan steps", indent=1)

    # Phase 10: Safety checks (warnings only - never block execution)
    spinner.set_task("Performing safety checks")
    progress("Phase 10: Performing preflight safety checks...")
    try:
        safety.perform_preflight_checks(generated_plan)
        progress("All preflight checks passed", indent=1, newline=False)
        progress_ok()
    except ValueError as err:
        progress(f"Warning: {err}", indent=1, newline=False)
        progress_skip("non-fatal")
        human_log.warning("Preflight safety warning: %s", err)
    except Exception as err:  # noqa: BLE001
        progress(f"Warning: {err}", indent=1, newline=False)
        progress_skip("non-fatal")
        human_log.warning("Unexpected preflight warning: %s", err)

    # Phase 11: Artifacts
    spinner.set_task("Writing plan artifacts")
    progress("Phase 11: Writing plan artifacts...")
    handle_plan_artifacts(args, generated_plan, inventory, human_log, mode)

    if mode == "diagnose":
        spinner.clear_task()
        spinner.stop_spinner_thread()
        progress("=" * 60)
        progress("Diagnostics complete - no actions executed")
        progress("=" * 60)
        human_log.info("Diagnostics complete; plan written and no actions executed.")
        return 0

    # Phase 12: User confirmation
    spinner.set_task("Awaiting confirmation")
    progress("Phase 12: Requesting user confirmation...")
    scrub_dry_run = bool(getattr(args, "dry_run", False))
    is_auto_all = mode == "auto-all"
    proceed = confirm.request_scrub_confirmation(
        dry_run=scrub_dry_run,
        force=bool(getattr(args, "force", False)) or is_auto_all,
    )
    if not proceed:
        spinner.clear_task()
        spinner.stop_spinner_thread()
        progress("User declined confirmation - aborting")
        human_log.info("Scrub cancelled by user confirmation prompt.")
        machine_log.info(
            "scrub.cancelled",
            extra={
                "event": "scrub.cancelled",
                "data": {"reason": "user_declined", "dry_run": scrub_dry_run},
            },
        )
        return 0
    progress("Confirmation received", indent=1)

    # Phase 13: Runtime guards (warnings only - never block execution)
    spinner.set_task("Enforcing runtime guards")
    progress("Phase 13: Enforcing runtime guards...")
    try:
        _enforce_runtime_guards(options, dry_run=scrub_dry_run)
        progress("Runtime guards passed", indent=1, newline=False)
        progress_ok()
    except ValueError as err:
        progress(f"Warning: {err}", indent=1, newline=False)
        progress_skip("non-fatal")
        human_log.warning("Runtime guard warning: %s", err)
    except Exception as err:  # noqa: BLE001
        progress(f"Warning: {err}", indent=1, newline=False)
        progress_skip("non-fatal")
        human_log.warning("Unexpected runtime warning: %s", err)

    # Phase 14: Plan execution
    spinner.set_task("Executing scrub plan")
    progress("=" * 60)
    progress(f"Phase 14: Executing plan ({'DRY RUN' if scrub_dry_run else 'LIVE'})...")
    progress("=" * 60)
    fatal_error: str | None = None
    try:
        scrub.execute_plan(generated_plan, dry_run=scrub_dry_run, start_time=_MAIN_START_TIME)
    except KeyboardInterrupt:
        progress("Execution interrupted by user", indent=1, newline=False)
        progress_skip("cancelled")
        human_log.warning("Plan execution cancelled by user")
        raise
    except (OSError, PermissionError) as err:
        # These are potentially fatal - system-level failures
        fatal_error = str(err)
        progress(f"Fatal error: {err}", indent=1, newline=False)
        progress_fail()
        human_log.error("Fatal execution error: %s", err)
    except Exception as err:  # noqa: BLE001
        # Log but don't treat as fatal - execution may have partially succeeded
        progress(f"Error during execution: {err}", indent=1, newline=False)
        progress_fail()
        human_log.exception("Error during plan execution (non-fatal)")

    # Stop spinner before final messages
    spinner.clear_task()
    spinner.stop_spinner_thread()

    progress("=" * 60)
    if fatal_error:
        progress(f"Execution failed after {get_elapsed_secs():.3f}s")
        progress("=" * 60)
        return 1
    progress(f"Execution complete in {get_elapsed_secs():.3f}s")
    progress("=" * 60)
    return 0


if __name__ == "__main__":  # pragma: no cover - for manual execution
    sys.exit(main())
