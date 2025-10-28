"""!
@brief Primary entry point for the Office Janitor CLI.
@details This module bootstraps the runtime environment described in
:mod:`spec.md`: argument parsing for detection and scrubbing modes, validating
administrative elevation, enabling Windows VT mode when available, and invoking
logging setup so future sub-systems can emit structured telemetry.
"""
from __future__ import annotations

import argparse
import ctypes
import datetime
import json
import logging
import os
import pathlib
import platform
import subprocess
import sys
from collections import deque
from typing import Iterable, List, Mapping, Optional

from . import (
    constants,
    detect,
    fs_tools,
    logging_ext,
    plan as plan_module,
    processes,
    safety,
    scrub,
    ui,
    tui,
    version,
)


def enable_vt_mode_if_possible() -> None:
    """!
    @brief Attempt to enable ANSI/VT processing on Windows consoles.
    @details Per the specification, the application should try to enable virtual
    terminal support so both the plain CLI and future TUI renderer can emit
    colorized output. Failures are silently ignored because the feature is
    optional.
    """

    if os.name != "nt":  # pragma: no cover - Windows behaviour only
        return

    try:
        from ctypes import wintypes

        kernel32 = ctypes.windll.kernel32  # type: ignore[attr-defined]
    except Exception:  # pragma: no cover - import/attribute errors on non-Windows
        return

    for std_handle in (-11, -12):  # STD_OUTPUT_HANDLE, STD_ERROR_HANDLE
        handle = kernel32.GetStdHandle(std_handle)
        if not handle:
            continue
        mode = wintypes.DWORD()
        if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
            kernel32.SetConsoleMode(handle, mode.value | 0x0004)


def ensure_admin_and_relaunch_if_needed() -> None:
    """!
    @brief Request elevation if the current process lacks administrative rights.
    """

    if os.name != "nt":
        return

    try:
        shell32 = ctypes.windll.shell32  # type: ignore[attr-defined]
    except Exception:  # pragma: no cover - non-Windows platforms
        return

    try:
        is_admin = bool(shell32.IsUserAnAdmin())
    except Exception:
        is_admin = False

    if is_admin:
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

    params = subprocess.list2cmdline(argv)
    result = shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
    if int(result) <= 32:
        raise SystemExit("Failed to request elevation via ShellExecuteW.")
    sys.exit(0)


def build_arg_parser() -> argparse.ArgumentParser:
    """!
    @brief Create the top-level argument parser with the specification's surface area.
    @details The parser wires in the mutually exclusive modes and shared options
    defined in the specification so later feature work can hook into the parsed
    values without changing the public CLI signature.
    """

    parser = argparse.ArgumentParser(prog="office-janitor", add_help=True)
    metadata = version.build_info()
    parser.add_argument(
        "-V",
        "--version",
        action="version",
        version=f"{metadata['version']} ({metadata['build']})",
    )

    modes = parser.add_mutually_exclusive_group()
    modes.add_argument("--auto-all", action="store_true", help="Run full detection and scrub.")
    modes.add_argument(
        "--target",
        metavar="VER",
        help="Target a specific Office version (2003-2024/365).",
    )
    modes.add_argument("--diagnose", action="store_true", help="Emit inventory and plan without changes.")
    modes.add_argument("--cleanup-only", action="store_true", help="Skip uninstalls; clean residue and licensing.")

    parser.add_argument("--include", metavar="COMPONENTS", help="Additional suites/apps to include.")
    parser.add_argument("--force", action="store_true", help="Relax certain guardrails when safe.")
    parser.add_argument(
        "--allow-unsupported-windows",
        action="store_true",
        help="Permit execution on Windows releases below the supported minimum.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Simulate actions without modifying the system.")
    parser.add_argument("--no-restore-point", action="store_true", help="Skip creating a restore point.")
    parser.add_argument("--no-license", action="store_true", help="Skip license cleanup steps.")
    parser.add_argument("--keep-templates", action="store_true", help="Preserve user templates like normal.dotm.")
    parser.add_argument("--plan", metavar="OUT", help="Write the computed action plan to a JSON file.")
    parser.add_argument("--logdir", metavar="DIR", help="Directory for human/JSONL log output.")
    parser.add_argument("--backup", metavar="DIR", help="Destination for registry/file backups.")
    parser.add_argument("--timeout", metavar="SEC", type=int, help="Per-step timeout in seconds.")
    parser.add_argument("--quiet", action="store_true", help="Minimal console output (errors only).")
    parser.add_argument("--json", action="store_true", help="Mirror structured events to stdout.")
    parser.add_argument("--tui", action="store_true", help="Force the interactive text UI mode.")
    parser.add_argument("--no-color", action="store_true", help="Disable ANSI color codes.")
    parser.add_argument("--tui-compact", action="store_true", help="Use a compact TUI layout for small consoles.")
    parser.add_argument(
        "--tui-refresh",
        metavar="MS",
        type=int,
        help="Refresh interval for the TUI renderer in milliseconds.",
    )
    return parser


def _resolve_log_directory(candidate: Optional[str]) -> pathlib.Path:
    """!
    @brief Determine the log directory path using specification defaults when unspecified.
    """

    if candidate:
        return pathlib.Path(candidate).expanduser().resolve()
    default_dir = fs_tools.get_default_log_directory()
    expanded = default_dir.expanduser()
    try:
        return expanded.resolve()
    except Exception:
        return expanded


def _bootstrap_logging(args: argparse.Namespace) -> tuple[logging.Logger, logging.Logger]:
    """!
    @brief Initialize human and machine loggers using :mod:`logging_ext` helpers.
    @returns A tuple of configured human and machine loggers.
    """

    logdir = _resolve_log_directory(getattr(args, "logdir", None))
    logdir.mkdir(parents=True, exist_ok=True)
    setattr(args, "logdir", str(logdir))
    human_logger, machine_logger = logging_ext.setup_logging(
        logdir,
        json_to_stdout=getattr(args, "json", False),
    )
    setattr(args, "human_logger", human_logger)
    setattr(args, "machine_logger", machine_logger)
    if getattr(args, "quiet", False):
        human_logger.setLevel(logging.ERROR)
    return human_logger, machine_logger


def main(argv: Optional[Iterable[str]] = None) -> int:
    """!
    @brief Entry point invoked by the shim and PyInstaller bundle.
    @returns Process exit code integer.
    """

    ensure_admin_and_relaunch_if_needed()
    enable_vt_mode_if_possible()
    parser = build_arg_parser()
    args = parser.parse_args(list(argv) if argv is not None else None)
    human_log, machine_log = _bootstrap_logging(args)
    if getattr(args, "quiet", False):
        human_log.setLevel(logging.ERROR)

    mode = _determine_mode(args)
    machine_log.info(
        "startup",
        extra={"event": "startup", "data": {"mode": mode, "dry_run": bool(getattr(args, "dry_run", False))}},
    )

    app_state = _build_app_state(args, human_log, machine_log)
    if mode == "interactive":
        if getattr(args, "tui", False):
            tui.run_tui(app_state)
        else:
            tui_candidate = _should_use_tui(args)
            if tui_candidate:
                tui.run_tui(app_state)
            else:
                ui.run_cli(app_state)
        return 0

    logdir_path = pathlib.Path(getattr(args, "logdir", _resolve_log_directory(None))).expanduser()
    inventory = _run_detection(machine_log, logdir_path)
    options = _collect_plan_options(args, mode)
    generated_plan = plan_module.build_plan(inventory, options)
    safety.perform_preflight_checks(generated_plan)
    _handle_plan_artifacts(args, generated_plan, inventory, human_log, mode)

    if mode == "diagnose":
        human_log.info("Diagnostics complete; plan written and no actions executed.")
        return 0

    scrub_dry_run = bool(getattr(args, "dry_run", False))
    _enforce_runtime_guards(options, dry_run=scrub_dry_run)
    scrub.execute_plan(generated_plan, dry_run=scrub_dry_run)
    return 0


def _determine_mode(args: argparse.Namespace) -> str:
    """!
    @brief Map parsed arguments to a simple textual mode identifier.
    """

    if getattr(args, "auto_all", False):
        return "auto-all"
    if getattr(args, "target", None):
        return f"target:{args.target}"
    if getattr(args, "diagnose", False):
        return "diagnose"
    if getattr(args, "cleanup_only", False):
        return "cleanup-only"
    return "interactive"


def _should_use_tui(args: argparse.Namespace) -> bool:
    """!
    @brief Determine whether the TUI should be launched automatically.
    @details The logic prefers the richer interface when the output stream
    supports ANSI escape codes and the caller did not explicitly disable color.
    """

    if getattr(args, "no_color", False):
        return False
    if getattr(sys.stdout, "isatty", None) and sys.stdout.isatty():
        return bool(os.environ.get("WT_SESSION") or os.environ.get("TERM", "").lower() != "dumb")
    return False


def _build_app_state(
    args: argparse.Namespace, human_log: logging.Logger, machine_log: logging.Logger
) -> Mapping[str, object]:
    """!
    @brief Assemble the dependency dictionary consumed by CLI/TUI front-ends.
    @details The mapping exposes callables for detection, planning, and
    execution so interactive interfaces can drive the same back-end flows as
    the non-interactive CLI code path.
    """

    ui_events: deque[dict[str, object]] = deque()

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

    def detector() -> dict:
        logdir_path = pathlib.Path(getattr(args, "logdir", _resolve_log_directory(None))).expanduser()
        return _run_detection(machine_log, logdir_path)

    def planner(inventory: Mapping[str, object], overrides: Optional[Mapping[str, object]] = None) -> list[dict]:
        mode = _determine_mode(args)
        merged = dict(_collect_plan_options(args, mode))
        if overrides:
            merged.update({key: overrides[key] for key in overrides})
        generated_plan = plan_module.build_plan(dict(inventory), merged)
        safety.perform_preflight_checks(generated_plan)
        return generated_plan

    def executor(plan_data: list[dict], overrides: Optional[Mapping[str, object]] = None) -> None:
        dry_run = bool(getattr(args, "dry_run", False))
        if overrides and "dry_run" in overrides:
            dry_run = bool(overrides["dry_run"])

        mode_override = _determine_mode(args)
        if overrides and overrides.get("mode"):
            mode_override = str(overrides["mode"])

        inventory_override = overrides.get("inventory") if overrides else None
        _handle_plan_artifacts(args, plan_data, inventory_override, human_log, mode_override)

        if mode_override == "diagnose":
            human_log.info("Diagnostics complete; plan written and no actions executed.")
            return

        guard_options = dict(_collect_plan_options(args, mode_override))
        guard_options["dry_run"] = dry_run
        if overrides:
            guard_options.update({key: overrides[key] for key in overrides})

        _enforce_runtime_guards(guard_options, dry_run=dry_run)
        scrub.execute_plan(plan_data, dry_run=dry_run)

    return {
        "args": args,
        "human_logger": human_log,
        "machine_logger": machine_log,
        "detector": detector,
        "planner": planner,
        "executor": executor,
        "event_queue": ui_events,
        "emit_event": emit_event,
    }


def _collect_plan_options(args: argparse.Namespace, mode: str) -> dict:
    """!
    @brief Translate parsed CLI arguments into planning options.
    """

    options = {
        "mode": mode,
        "dry_run": bool(getattr(args, "dry_run", False)),
        "force": bool(getattr(args, "force", False)),
        "include": getattr(args, "include", None),
        "target": getattr(args, "target", None),
        "diagnose": bool(getattr(args, "diagnose", False)),
        "cleanup_only": bool(getattr(args, "cleanup_only", False)),
        "auto_all": bool(getattr(args, "auto_all", False)),
        "allow_unsupported_windows": bool(
            getattr(args, "allow_unsupported_windows", False)
        ),
        "no_license": bool(getattr(args, "no_license", False)),
        "keep_templates": bool(getattr(args, "keep_templates", False)),
        "timeout": getattr(args, "timeout", None),
        "backup": getattr(args, "backup", None),
        "create_restore_point": not bool(getattr(args, "no_restore_point", False)),
    }
    return options


def _run_detection(
    machine_log: logging.Logger, log_directory: pathlib.Path | str | None = None
) -> dict:
    """!
    @brief Execute inventory gathering, persist artifacts, and emit telemetry.
    """

    inventory = detect.gather_office_inventory()
    if log_directory is None:
        logdir_path = _resolve_log_directory(None)
    else:
        logdir_path = pathlib.Path(log_directory).expanduser()

    inventory_path: pathlib.Path | None = None
    try:
        logdir_path.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%d-%H%M%S")
        inventory_path = logdir_path / f"inventory-{timestamp}.json"
        inventory_path.write_text(
            json.dumps(inventory, indent=2, sort_keys=True, default=str),
            encoding="utf-8",
        )
    except OSError as exc:
        machine_log.warning(
            "inventory_write_failed",
            extra={
                "event": "inventory_write_failed",
                "error": repr(exc),
                "logdir": str(logdir_path),
            },
        )

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


def _handle_plan_artifacts(
    args: argparse.Namespace,
    plan_data: Iterable[Mapping[str, object]],
    inventory: Optional[Mapping[str, object]],
    human_log: logging.Logger,
    mode: str,
) -> None:
    """!
    @brief Persist plan diagnostics and backups as requested via CLI flags.
    """

    plan_steps = list(plan_data)
    logdir = pathlib.Path(getattr(args, "logdir", _resolve_log_directory(None))).expanduser()
    logdir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%d-%H%M%S")
    resolved_backup: Optional[pathlib.Path] = None

    backup_dir = getattr(args, "backup", None)
    if backup_dir:
        destination = pathlib.Path(backup_dir).expanduser().resolve()
        destination.mkdir(parents=True, exist_ok=True)
        if inventory is not None:
            (destination / "inventory.json").write_text(
                json.dumps(inventory, indent=2, sort_keys=True),
                encoding="utf-8",
            )
        resolved_backup = destination
    else:
        resolved_backup = logdir / f"registry-backup-{timestamp}"

    if mode == "diagnose" and inventory is not None and not backup_dir:
        inventory_path = logdir / "diagnostics-inventory.json"
        inventory_path.write_text(
            json.dumps(inventory, indent=2, sort_keys=True),
            encoding="utf-8",
        )
        human_log.info("Wrote diagnostics inventory to %s", inventory_path)

    if plan_steps:
        context_step = next((step for step in plan_steps if step.get("category") == "context"), None)
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
            for step in plan_steps:
                if step.get("category") != "registry-cleanup":
                    continue
                registry_metadata = dict(step.get("metadata", {}))
                registry_metadata.setdefault("backup_destination", str(resolved_backup))
                registry_metadata.setdefault("log_directory", str(logdir))
                step["metadata"] = registry_metadata

    serialized_plan = json.dumps(plan_steps, indent=2, sort_keys=True)
    primary_plan_path = logdir / f"plan-{timestamp}.json"
    primary_plan_path.write_text(serialized_plan, encoding="utf-8")
    human_log.info("Wrote plan to %s", primary_plan_path)

    additional_plan_targets: List[pathlib.Path] = []
    if getattr(args, "plan", None):
        additional_plan_targets.append(pathlib.Path(args.plan).expanduser().resolve())
    elif mode == "diagnose":
        additional_plan_targets.append(logdir / "diagnostics-plan.json")

    for target in additional_plan_targets:
        if target == primary_plan_path:
            continue
        target.write_text(serialized_plan, encoding="utf-8")
        human_log.info("Wrote plan to %s", target)

    if backup_dir and resolved_backup is not None:
        (resolved_backup / "plan.json").write_text(serialized_plan, encoding="utf-8")
        human_log.info("Wrote backup artifacts to %s", resolved_backup)


def _enforce_runtime_guards(options: Mapping[str, object], *, dry_run: bool) -> None:
    """!
    @brief Evaluate runtime safety prerequisites prior to executing the scrubber.
    @details Gathers host telemetry and forwards it to
    :func:`safety.evaluate_runtime_environment` so operating system, process, and
    restore point guards are enforced consistently across CLI entry points.
    """

    system, release = _detect_operating_system()
    require_restore_point = bool(options.get("create_restore_point", False))
    restore_point_available = True
    if require_restore_point and not dry_run:
        restore_point_available = _restore_points_available()

    safety.evaluate_runtime_environment(
        is_admin=_current_process_is_admin(),
        os_system=system,
        os_release=release,
        blocking_processes=_discover_blocking_processes(),
        dry_run=dry_run,
        require_restore_point=require_restore_point,
        restore_point_available=restore_point_available,
        force=bool(options.get("force", False)),
        allow_unsupported_windows=bool(options.get("allow_unsupported_windows", False)),
    )


def _detect_operating_system() -> tuple[str, str]:
    """!
    @brief Collect the current operating system identifier and release version.
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
    """

    patterns = list(constants.DEFAULT_OFFICE_PROCESSES) + list(constants.OFFICE_PROCESS_PATTERNS)
    try:
        return processes.enumerate_processes(patterns)
    except Exception:
        return []


def _current_process_is_admin() -> bool:
    """!
    @brief Determine whether the current interpreter is running with elevated privileges.
    """

    if os.name == "nt":
        try:
            shell32 = ctypes.windll.shell32  # type: ignore[attr-defined]
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

    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=15,
            check=False,
        )
    except FileNotFoundError:
        return False
    except Exception:
        return False

    return result.returncode == 0


if __name__ == "__main__":  # pragma: no cover - for manual execution
    sys.exit(main())
