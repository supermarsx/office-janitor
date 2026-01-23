"""!
@file main_state.py
@brief Application state and plan utilities for Office Janitor.
@details Handles app state construction, detection, artifacts, and runtime guards.
"""

from __future__ import annotations

import ctypes
import datetime
import json
import logging
import os
import pathlib
import platform
from collections.abc import Iterable, Mapping
from typing import TYPE_CHECKING

from . import (
    confirm,
    constants,
    detect,
    exec_utils,
    fs_tools,
    logging_ext,
    processes,
    safety,
    scrub,
    spinner,
)
from . import plan as plan_module
from .app_state import AppState, new_event_queue
from .main_plan import collect_plan_options, load_config_file  # noqa: F401
from .main_progress import (
    get_elapsed_secs,
    get_progress_lock,
    progress,
    progress_fail,
    progress_ok,
    progress_skip,
)

if TYPE_CHECKING:
    import argparse

__all__ = [
    "build_app_state",
    "collect_plan_options",
    "run_detection",
    "handle_plan_artifacts",
    "enforce_runtime_guards",
    "determine_mode",
    "should_use_tui",
    "resolve_log_directory",
    "load_config_file",
]


def resolve_log_directory(candidate: str | None) -> pathlib.Path:
    """!
    @brief Determine the log directory path using specification defaults when unspecified.
    @param candidate Path specified by user, or None for default.
    @returns Resolved path to log directory.
    """
    if candidate:
        return pathlib.Path(candidate).expanduser().resolve()
    default_dir = fs_tools.get_default_log_directory()
    expanded = default_dir.expanduser()
    try:
        return expanded.resolve()
    except Exception:
        return expanded


def determine_mode(args: argparse.Namespace) -> str:
    """!
    @brief Map parsed arguments to a simple textual mode identifier.
    @param args Parsed command-line arguments.
    @returns Mode string identifying the operation type.
    @details Supports both new subcommand syntax and legacy flags for
    backward compatibility.

    Subcommand modes:
    - install: ODT installation
    - repair: Repair operations (quick, full, odt, c2r)
    - remove: Uninstall and scrub

    Legacy modes (backward compatibility):
    - auto-all, target:VER, diagnose, cleanup-only
    - auto-repair, repair-odt, repair-c2r, repair:TYPE
    - oem-config:NAME
    """
    # Handle new subcommand-based syntax
    command = getattr(args, "command", None)

    if command == "diagnose":
        return "diagnose"

    if command == "install":
        # Check for author aliases
        if getattr(args, "goobler", False):
            return "install:goobler"
        if getattr(args, "pupa", False):
            return "install:pupa"
        # Check for preset-based installation
        preset = getattr(args, "odt_preset", None)
        if preset:
            return f"install:preset:{preset}"
        # Check for ODT build/download
        if getattr(args, "odt_output", None):
            return "install:build"
        if getattr(args, "odt_download", None):
            return "install:download"
        # Check for list commands
        if getattr(args, "odt_list_presets", False):
            return "install:list-presets"
        if getattr(args, "odt_list_products", False):
            return "install:list-products"
        if getattr(args, "odt_list_channels", False):
            return "install:list-channels"
        if getattr(args, "odt_list_languages", False):
            return "install:list-languages"
        return "install:interactive"

    if command == "repair":
        # Check repair type
        repair_type = getattr(args, "repair_type", None)
        if repair_type == "quick":
            return "repair:quick"
        if repair_type == "full":
            return "repair:full"
        if getattr(args, "repair_odt", False):
            return "repair-odt"
        if getattr(args, "repair_c2r", False):
            return "repair-c2r"
        # Default to auto-repair if no specific type
        return "auto-repair"

    if command == "remove":
        # Check target selection
        target = getattr(args, "target", None)
        if target:
            return f"target:{target}"
        # Check uninstall method
        method = getattr(args, "uninstall_method", None)
        if method == "msi":
            return "remove:msi-only"
        if method == "c2r":
            return "remove:c2r-only"
        # Default to full removal
        return "auto-all"

    # ---------------------------------------------------------------------------
    # Legacy flag handling (backward compatibility)
    # ---------------------------------------------------------------------------
    if getattr(args, "auto_all", False):
        return "auto-all"
    if getattr(args, "target", None):
        return f"target:{args.target}"
    if getattr(args, "diagnose", False):
        return "diagnose"
    if getattr(args, "cleanup_only", False):
        return "cleanup-only"
    if getattr(args, "auto_repair", False):
        return "auto-repair"
    if getattr(args, "repair_odt", False):
        return "repair-odt"
    if getattr(args, "repair_c2r", False):
        return "repair-c2r"
    if getattr(args, "repair", None):
        return f"repair:{args.repair}"
    if getattr(args, "repair_config", None):
        return "repair:config"
    if getattr(args, "oem_config", None):
        return f"oem-config:{args.oem_config}"
    # Author aliases (legacy style)
    if getattr(args, "goobler", False):
        return "install:goobler"
    if getattr(args, "pupa", False):
        return "install:pupa"
    # ODT operations (legacy style)
    if getattr(args, "odt_install", False):
        preset = getattr(args, "odt_preset", None)
        if preset:
            return f"install:preset:{preset}"
        return "install:odt"
    if getattr(args, "odt_build", False):
        return "install:build"
    return "interactive"


def should_use_tui(args: argparse.Namespace) -> bool:
    """!
    @brief Determine whether the TUI should be launched automatically.
    @details The logic prefers the richer interface when the output stream
    supports ANSI escape codes and the caller did not explicitly disable color.
    @param args Parsed command-line arguments.
    @returns True if TUI should be used.
    """
    import sys

    if getattr(args, "no_color", False):
        return False
    if getattr(sys.stdout, "isatty", None) and sys.stdout.isatty():
        return bool(os.environ.get("WT_SESSION") or os.environ.get("TERM", "").lower() != "dumb")
    return False


def build_app_state(
    args: argparse.Namespace,
    human_log: logging.Logger,
    machine_log: logging.Logger,
    *,
    start_time: float | None = None,
) -> AppState:
    """!
    @brief Assemble the dependency dictionary consumed by CLI/TUI front-ends.
    @details The mapping exposes callables for detection, planning, and
    execution so interactive interfaces can drive the same back-end flows as
    the non-interactive CLI code path.
    @param args Parsed command-line arguments.
    @param human_log Human-readable logger.
    @param machine_log Machine-readable (JSONL) logger.
    @param start_time Optional start time for progress tracking.
    @returns AppState dictionary with all dependencies.
    """
    from .main_progress import get_main_start_time

    ui_events = new_event_queue()

    def emit_event(event: str, *, message: str | None = None, **payload: object) -> None:
        """!
        @brief Publish progress events for interactive front-ends.
        @details Events are queued for consumption by the CLI/TUI layers so they
        can surface background activity without polling implementation details.
        """
        record: dict[str, object] = {"event": event}
        if message is not None:
            record["message"] = message
        if payload:
            record["data"] = dict(payload)
        ui_events.append(record)

    logging_ext.register_ui_event_sink(emitter=emit_event, queue=ui_events)

    def detector() -> dict[str, object]:
        logdir_path = pathlib.Path(
            getattr(args, "logdir", resolve_log_directory(None))
        ).expanduser()
        return run_detection(
            machine_log,
            logdir_path,
            limited_user=bool(getattr(args, "limited_user", False)),
        )

    def planner(
        inventory: Mapping[str, object], overrides: Mapping[str, object] | None = None
    ) -> list[dict[str, object]]:
        mode = determine_mode(args)
        merged = dict(collect_plan_options(args, mode))
        if overrides:
            merged.update({key: overrides[key] for key in overrides})
        generated_plan = plan_module.build_plan(dict(inventory), merged)
        safety.perform_preflight_checks(generated_plan)
        return generated_plan

    def executor(
        plan_data: list[dict[str, object]], overrides: Mapping[str, object] | None = None
    ) -> bool:
        dry_run = bool(getattr(args, "dry_run", False))
        if overrides and "dry_run" in overrides:
            dry_run = bool(overrides["dry_run"])

        mode_override = determine_mode(args)
        if overrides and overrides.get("mode"):
            mode_override = str(overrides["mode"])

        inventory_override = overrides.get("inventory") if overrides else None
        handle_plan_artifacts(args, plan_data, inventory_override, human_log, mode_override)

        if mode_override == "diagnose":
            human_log.info("Diagnostics complete; plan written and no actions executed.")
            return True

        guard_options = dict(collect_plan_options(args, mode_override))
        guard_options["dry_run"] = dry_run
        if overrides:
            guard_options.update({key: overrides[key] for key in overrides})

        force_override = bool(getattr(args, "force", False))
        if overrides and "force" in overrides:
            force_override = bool(overrides["force"])

        confirmed = bool(overrides.get("confirmed")) if overrides else False
        if not confirmed:
            proceed = confirm.request_scrub_confirmation(
                dry_run=dry_run,
                force=force_override,
                input_func=overrides.get("input_func") if overrides else None,
                interactive=overrides.get("interactive") if overrides else None,
            )
            if not proceed:
                if human_log:
                    human_log.info("Scrub cancelled by user confirmation prompt.")
                if machine_log:
                    machine_log.info(
                        "scrub.cancelled",
                        extra={
                            "event": "scrub.cancelled",
                            "data": {"reason": "user_declined", "dry_run": dry_run},
                        },
                    )
                return False

        enforce_runtime_guards(guard_options, dry_run=dry_run)
        main_start = start_time if start_time is not None else get_main_start_time()
        scrub.execute_plan(plan_data, dry_run=dry_run, start_time=main_start)
        return True

    app_state: AppState = {
        "args": args,
        "human_logger": human_log,
        "machine_logger": machine_log,
        "detector": detector,
        "planner": planner,
        "executor": executor,
        "event_queue": ui_events,
        "emit_event": emit_event,
        "confirm": confirm.request_scrub_confirmation,
    }
    return app_state


def run_detection(
    machine_log: logging.Logger,
    log_directory: pathlib.Path | str | None = None,
    *,
    limited_user: bool | None = None,
) -> dict[str, object]:
    """!
    @brief Execute inventory gathering, persist artifacts, and emit telemetry.
    @param machine_log Machine-readable logger for telemetry.
    @param log_directory Directory to write inventory files.
    @param limited_user Whether to run under limited user token.
    @returns Dictionary containing the inventory.
    """
    progress("Starting inventory scan...", indent=1)

    if limited_user:
        machine_log.info("Detection requested under limited user token.")
        progress("Running under limited user token", indent=2)

    # Use detailed progress callback for inventory gathering.
    # Since detect.py calls this from multiple threads concurrently, we use
    # complete single-line output for thread-safety (no pending line pattern).
    _progress_lock = get_progress_lock()

    def progress_callback(phase: str, status: str = "start") -> None:
        prefix = "      "  # indent=3 equivalent
        timestamp = f"[{get_elapsed_secs():12.6f}]"
        spinner.pause_for_output()
        try:
            if status == "start":
                # Print complete line with "..." to indicate in-progress
                with _progress_lock:
                    print(f"{timestamp} {prefix}{phase}...", flush=True)
            elif status == "ok":
                with _progress_lock:
                    print(
                        f"{timestamp} {prefix}{phase} [  \033[32mOK\033[0m  ]",
                        flush=True,
                    )
            elif status == "skip":
                with _progress_lock:
                    print(
                        f"{timestamp} {prefix}{phase} [ \033[33mSKIP\033[0m ]",
                        flush=True,
                    )
            elif status == "fail":
                with _progress_lock:
                    print(
                        f"{timestamp} {prefix}{phase} [\033[31mFAILED\033[0m]",
                        flush=True,
                    )
        finally:
            spinner.resume_after_output()

    progress("Gathering Office inventory...", indent=2)
    try:
        if limited_user:
            inventory = detect.gather_office_inventory(
                limited_user=True, progress_callback=progress_callback
            )
        else:
            inventory = detect.gather_office_inventory(progress_callback=progress_callback)
        progress("Inventory collection complete", indent=2, newline=False)
        progress_ok()
    except KeyboardInterrupt:
        print(flush=True)  # Newline after partial output
        progress("Inventory collection interrupted", indent=2, newline=False)
        progress_skip("user cancelled")
        raise

    if log_directory is None:
        logdir_path = resolve_log_directory(None)
    else:
        logdir_path = pathlib.Path(log_directory).expanduser()

    progress(f"Log directory: {logdir_path}", indent=2)

    inventory_path: pathlib.Path | None = None
    try:
        logdir_path.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%d-%H%M%S")
        inventory_path = logdir_path / f"inventory-{timestamp}.json"
        progress(f"Writing inventory to {inventory_path.name}...", indent=2, newline=False)
        inventory_path.write_text(
            json.dumps(inventory, indent=2, sort_keys=True, default=str),
            encoding="utf-8",
        )
        progress_ok()
    except OSError as exc:
        progress_fail(str(exc))
        machine_log.warning(
            "inventory_write_failed",
            extra={
                "event": "inventory_write_failed",
                "error": repr(exc),
                "logdir": str(logdir_path),
            },
        )

    # Log inventory summary
    progress("Inventory summary:", indent=2)
    for key, value in inventory.items():
        count = len(value) if hasattr(value, "__len__") else len(list(value))
        if count > 0:
            progress(f"{key}: {count} items", indent=3)

    machine_log.info(
        "inventory",
        extra={
            "event": "inventory",
            "counts": {
                key: len(value) if hasattr(value, "__len__") else len(list(value))
                for key, value in inventory.items()
            },
            **({"artifact": str(inventory_path)} if inventory_path is not None else {}),
        },
    )
    return inventory


def handle_plan_artifacts(
    args: argparse.Namespace,
    plan_data: Iterable[Mapping[str, object]],
    inventory: Mapping[str, object] | None,
    human_log: logging.Logger,
    mode: str,
) -> None:
    """!
    @brief Persist plan diagnostics and backups as requested via CLI flags.
    @param args Parsed command-line arguments.
    @param plan_data The execution plan steps.
    @param inventory The detected inventory (or None).
    @param human_log Human-readable logger.
    @param mode Operation mode string.
    """
    progress("Processing plan artifacts...", indent=1)

    plan_steps = list(plan_data)
    logdir = pathlib.Path(getattr(args, "logdir", resolve_log_directory(None))).expanduser()
    logdir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%d-%H%M%S")
    resolved_backup: pathlib.Path | None = None

    backup_dir = getattr(args, "backup", None)
    if backup_dir:
        progress(f"Setting up backup directory: {backup_dir}", indent=2)
        destination = pathlib.Path(backup_dir).expanduser().resolve()
        destination.mkdir(parents=True, exist_ok=True)
        if inventory is not None:
            progress("Writing inventory to backup...", indent=3, newline=False)
            (destination / "inventory.json").write_text(
                json.dumps(inventory, indent=2, sort_keys=True),
                encoding="utf-8",
            )
            progress_ok()
        resolved_backup = destination
    else:
        resolved_backup = logdir / f"registry-backup-{timestamp}"

    if mode == "diagnose" and inventory is not None and not backup_dir:
        inventory_path = logdir / "diagnostics-inventory.json"
        progress(f"Writing diagnostics inventory: {inventory_path.name}", indent=2)
        inventory_path.write_text(
            json.dumps(inventory, indent=2, sort_keys=True),
            encoding="utf-8",
        )
        human_log.info("Wrote diagnostics inventory to %s", inventory_path)

    if plan_steps:
        progress(f"Enriching {len(plan_steps)} plan steps with metadata...", indent=2)
        context_step = next(
            (step for step in plan_steps if step.get("category") == "context"), None
        )
        if context_step is not None:
            metadata = dict(context_step.get("metadata", {}))
            options = dict(metadata.get("options", {}))
            metadata["log_directory"] = str(logdir)
            options["log_directory"] = str(logdir)
            options["logdir"] = str(logdir)
            if resolved_backup is not None:
                metadata["backup_destination"] = str(resolved_backup)
                options["backup"] = str(resolved_backup)
            metadata["options"] = options
            context_step["metadata"] = metadata

        if resolved_backup is not None:
            registry_steps = sum(1 for s in plan_steps if s.get("category") == "registry-cleanup")
            if registry_steps > 0:
                progress(f"Configuring backup for {registry_steps} registry steps", indent=2)
            for step in plan_steps:
                if step.get("category") != "registry-cleanup":
                    continue
                registry_metadata = dict(step.get("metadata", {}))
                registry_metadata.setdefault("backup_destination", str(resolved_backup))
                registry_metadata.setdefault("log_directory", str(logdir))
                step["metadata"] = registry_metadata

    serialized_plan = json.dumps(plan_steps, indent=2, sort_keys=True)
    primary_plan_path = logdir / f"plan-{timestamp}.json"
    progress(f"Writing primary plan: {primary_plan_path.name}", indent=2)
    primary_plan_path.write_text(serialized_plan, encoding="utf-8")
    human_log.info("Wrote plan to %s", primary_plan_path)

    additional_plan_targets: list[pathlib.Path] = []
    if getattr(args, "plan", None):
        additional_plan_targets.append(pathlib.Path(args.plan).expanduser().resolve())
    elif mode == "diagnose":
        additional_plan_targets.append(logdir / "diagnostics-plan.json")

    for target in additional_plan_targets:
        if target == primary_plan_path:
            continue
        progress(f"Writing additional plan: {target.name}", indent=2)
        target.write_text(serialized_plan, encoding="utf-8")
        human_log.info("Wrote plan to %s", target)

    if backup_dir and resolved_backup is not None:
        progress(f"Writing plan to backup: {resolved_backup}", indent=2)
        (resolved_backup / "plan.json").write_text(serialized_plan, encoding="utf-8")
        human_log.info("Wrote backup artifacts to %s", resolved_backup)

    progress("Artifact processing complete", indent=1)


def enforce_runtime_guards(options: Mapping[str, object], *, dry_run: bool) -> None:
    """!
    @brief Evaluate runtime safety prerequisites prior to executing the scrubber.
    @details Gathers host telemetry and forwards it to
    :func:`safety.evaluate_runtime_environment` so operating system, process, and
    restore point guards are enforced consistently across CLI entry points.
    @param options Planning options dictionary.
    @param dry_run Whether this is a dry-run execution.
    """
    progress("Gathering runtime environment info...", indent=1)

    progress("Detecting operating system...", indent=2, newline=False)
    system, release = _detect_operating_system()
    progress_ok(f"{system} {release}")

    require_restore_point = bool(options.get("create_restore_point", False))
    restore_point_available = True
    if require_restore_point and not dry_run:
        progress("Checking restore point availability...", indent=2, newline=False)
        restore_point_available = _restore_points_available()
        if restore_point_available:
            progress_ok()
        else:
            progress_fail("not available")

    progress("Checking admin privileges...", indent=2, newline=False)
    is_admin = _current_process_is_admin()
    if is_admin:
        progress_ok()
    else:
        progress_fail()

    progress("Scanning for blocking processes...", indent=2, newline=False)
    blocking = _discover_blocking_processes()
    if blocking:
        progress_fail(f"{len(blocking)} found")
        for proc in blocking[:5]:  # Show first 5
            progress(f"- {proc}", indent=3)
        if len(blocking) > 5:
            progress(f"... and {len(blocking) - 5} more", indent=3)
    else:
        progress_ok("none")

    progress("Evaluating safety constraints...", indent=2, newline=False)
    safety.evaluate_runtime_environment(
        is_admin=is_admin,
        os_system=system,
        os_release=release,
        blocking_processes=blocking,
        dry_run=dry_run,
        require_restore_point=require_restore_point,
        restore_point_available=restore_point_available,
        force=bool(options.get("force", False)),
        allow_unsupported_windows=bool(options.get("allow_unsupported_windows", False)),
        minimum_free_space_bytes=options.get("minimum_free_space_bytes"),
        disk_usage_root=options.get("disk_usage_root")
        or options.get("free_space_root")
        or options.get("system_drive"),
    )
    progress_ok()


def _detect_operating_system() -> tuple[str, str]:
    """!
    @brief Collect the current operating system identifier and release version.
    @returns Tuple of (system, release) strings.
    """
    try:
        system = platform.system()
    except Exception:
        system = ""

    release = ""
    try:
        release = platform.version()
    except Exception:
        release = ""

    if not release:
        try:
            release = platform.release()
        except Exception:
            release = ""

    return system, release


def _discover_blocking_processes() -> list[str]:
    """!
    @brief Enumerate Office-related processes that may block destructive actions.
    @returns List of blocking process names.
    """
    patterns = list(constants.DEFAULT_OFFICE_PROCESSES) + list(constants.OFFICE_PROCESS_PATTERNS)
    try:
        return processes.enumerate_processes(patterns)
    except Exception:
        return []


def _current_process_is_admin() -> bool:
    """!
    @brief Determine whether the current interpreter is running with elevated privileges.
    @returns True if running as admin/root.
    """
    if os.name == "nt":
        try:
            shell32 = ctypes.windll.shell32
            return bool(shell32.IsUserAnAdmin())
        except Exception:
            return False

    geteuid = getattr(os, "geteuid", None)
    if callable(geteuid):
        try:
            return bool(geteuid() == 0)
        except Exception:
            return False

    return False


def _restore_points_available() -> bool:
    """!
    @brief Detect whether system restore points are currently available.
    @returns True if restore points are available.
    """
    if os.name != "nt":
        return False

    script = "\n".join(
        (
            "Try {",
            "  Get-ComputerRestorePoint -ErrorAction Stop | Select-Object -First 1 | Out-String",
            "  Exit 0",
            " } Catch { Exit 1 }",
        )
    )

    command = [
        "powershell.exe",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        script,
    ]

    result = exec_utils.run_command(
        command,
        event="restore_point_probe",
        timeout=15,
    )

    return result.returncode == 0 and not result.error
