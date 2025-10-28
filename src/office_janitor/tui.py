"""!
@brief Text-based user interface (TUI) engine.
@details Implements the ANSI/VT driven interface with panes, widgets, and event
queue plumbing described in the project specification. The implementation keeps
dependencies to the standard library only while providing a co-operative event
loop that drains orchestrator progress events and handles keyboard commands.
"""
from __future__ import annotations

import os
import sys
import time
from collections import deque
from typing import Callable, Deque, List, Mapping, MutableMapping, Optional

from . import plan as plan_module
from . import version


class OfficeJanitorTUI:
    """!
    @brief Controller coordinating rendering, events, and orchestrator calls.
    """

    def __init__(self, app_state: Mapping[str, object]) -> None:
        self.app_state: MutableMapping[str, object] = dict(app_state)
        self.human_logger = self.app_state.get("human_logger")
        self.machine_logger = self.app_state.get("machine_logger")
        self.detector: Callable[[], Mapping[str, object]] = self.app_state[
            "detector"
        ]  # type: ignore[assignment]
        self.planner: Callable[[Mapping[str, object], Mapping[str, object] | None], list[dict]] = self.app_state[
            "planner"
        ]  # type: ignore[assignment]
        self.executor: Callable[[list[dict], Mapping[str, object] | None], None] = self.app_state[
            "executor"
        ]  # type: ignore[assignment]
        queue_obj = self.app_state.get("event_queue")
        if isinstance(queue_obj, deque):
            self.event_queue: Deque[dict[str, object]] = queue_obj
        else:
            self.event_queue = deque()
            self.app_state["event_queue"] = self.event_queue
        self.emit_event = self.app_state.get("emit_event")
        self.last_inventory: Optional[Mapping[str, object]] = None
        self.last_plan: Optional[list[dict]] = None
        self.status_lines: List[str] = []
        self.progress_message = "Ready"
        self._key_reader: Optional[Callable[[], str]] = self.app_state.get(
            "key_reader"
        )  # type: ignore[assignment]
        args = self.app_state.get("args")
        refresh_ms = getattr(args, "tui_refresh", 120) if args is not None else 120
        try:
            refresh_value = float(refresh_ms) / 1000.0
        except Exception:
            refresh_value = 0.12
        self.refresh_interval = 0.05 if refresh_value <= 0 else refresh_value
        self.compact_layout = bool(getattr(args, "tui_compact", False)) if args is not None else False
        self._running = True

    def run(self) -> None:
        """!
        @brief Enter the TUI event loop.
        """

        args = self.app_state.get("args")
        if getattr(args, "quiet", False) or getattr(args, "json", False):
            if self.human_logger:
                self.human_logger.info(
                    "Interactive TUI suppressed because quiet/json output mode was requested.",
                )
            self._notify("tui.suppressed", "TUI launch suppressed by CLI flags.")
            return

        if getattr(args, "no_color", False) or not _supports_ansi():
            from . import ui

            self._notify("tui.fallback", "Falling back to CLI menu (ANSI unavailable).")
            ui.run_cli(self.app_state)
            return

        self._notify("tui.start", "Interactive TUI started.")
        self._render()

        while self._running:
            if self._drain_events():
                self._render()

            command = self._read_command()
            if not command:
                time.sleep(self.refresh_interval)
                continue

            if command in {"q", "Q", "\u001b"}:
                self.progress_message = "Exiting..."
                self._notify("tui.exit", "User requested exit from TUI.")
                self._render()
                break
            if command in {"d", "D"}:
                self._handle_detect()
            elif command in {"p", "P"}:
                self._handle_plan()
            elif command in {"r", "R", "\r"}:
                self._handle_run()
            elif command in {"l", "L"}:
                self._handle_logs()
            elif command in {"a", "A"}:
                self._handle_mode(
                    "auto-all",
                    {"mode": "auto-all", "auto_all": True},
                    friendly="auto scrub",
                )
            elif command in {"t", "T"}:
                self._handle_targeted()
            elif command in {"c", "C"}:
                self._handle_mode(
                    "cleanup-only",
                    {"mode": "cleanup-only", "cleanup_only": True},
                    friendly="cleanup",
                )
            elif command in {"g", "G"}:
                self._handle_mode(
                    "diagnose",
                    {"mode": "diagnose", "diagnose": True},
                    friendly="diagnostics",
                )
            elif command in {"s", "S"}:
                self._handle_settings()
            else:
                self._notify("tui.unknown", f"Unknown command: {command!r}", level="warning")
            self._render()

    def _render(self) -> None:
        width = 80 if self.compact_layout else 96
        left_width = 32 if self.compact_layout else 36
        _clear_screen()
        metadata = version.build_info()
        header = f"Office Janitor {metadata['version']} — {self.progress_message}"
        sys.stdout.write(header[:width] + "\n")
        sys.stdout.write(_divider(width) + "\n")

        left_lines = [
            "[D] Detect inventory",
            "[P] Build plan",
            "[R] Run current plan",
            "[A] Auto scrub everything",
            "[T] Targeted scrub",
            "[C] Cleanup only",
            "[G] Diagnostics only",
            "[L] Log info",
            "[S] Settings",
            "[Q] Quit",
            "",
            "Status log:",
        ] + self.status_lines[-(12 if self.compact_layout else 18) :]

        inventory_lines = (
            _format_inventory(self.last_inventory)
            if self.last_inventory is not None
            else ["No inventory collected"]
        )
        plan_lines = _format_plan(self.last_plan)
        right_lines = ["Inventory summary:"] + inventory_lines + ["", "Plan summary:"] + plan_lines

        max_lines = max(len(left_lines), len(right_lines))
        for index in range(max_lines):
            left_text = left_lines[index] if index < len(left_lines) else ""
            right_text = right_lines[index] if index < len(right_lines) else ""
            sys.stdout.write(f"{left_text.ljust(left_width)} {right_text[: width - left_width - 1]}\n")

        sys.stdout.write(_divider(width) + "\n")
        sys.stdout.write(
            "Commands: d=Detect p=Plan r=Run a=Auto t=Target c=Cleanup g=Diagnose l=Logs s=Settings q=Quit\n"
        )
        sys.stdout.flush()

    def _read_command(self) -> str:
        reader = self._key_reader or _default_key_reader
        try:
            command = reader()
        except StopIteration:
            self._running = False
            return "q"
        except Exception:
            return ""
        if isinstance(command, str):
            return command.strip()[:1]
        return ""

    def _handle_detect(self) -> None:
        self.progress_message = "Detecting inventory..."
        self._notify("detect.start", "Starting detection run from TUI.")
        self._render()
        try:
            inventory = self.detector()
        except Exception as exc:  # pragma: no cover - defensive logging
            message = f"Detection failed: {exc}"
            self._notify("detect.error", message, level="error")
            self.progress_message = "Detection failed"
            return

        self.last_inventory = inventory
        summary = _summarize_inventory(inventory)
        self._notify("detect.complete", "Inventory captured.", inventory=summary)
        for line in _format_inventory(inventory):
            self._append_status(f"  {line}")
        self.progress_message = "Inventory ready"

    def _handle_plan(self) -> None:
        if self.last_inventory is None:
            self._handle_detect()
            if self.last_inventory is None:
                return

        self.progress_message = "Planning actions..."
        self._notify("plan.start", "Building plan from TUI.")
        self._render()
        try:
            plan_data = self.planner(self.last_inventory, None)
        except Exception as exc:  # pragma: no cover - defensive logging
            message = f"Plan failed: {exc}"
            self._notify("plan.error", message, level="error")
            self.progress_message = "Plan failed"
            return

        self.last_plan = plan_data
        summary = plan_module.summarize_plan(plan_data)
        self._notify("plan.complete", "Plan ready for review.", summary=summary)
        for line in _format_plan(plan_data)[:6]:
            self._append_status(f"  {line}")
        self.progress_message = "Plan ready"

    def _handle_run(self) -> None:
        if self.last_plan is None:
            self._handle_plan()
            if self.last_plan is None:
                return

        self.progress_message = "Executing plan..."
        self._notify("execution.start", "Executing plan from TUI.")
        self._render()
        try:
            _spinner(0.2, "Preparing")
            self.executor(self.last_plan, None)
        except Exception as exc:  # pragma: no cover - defensive logging
            message = f"Execution failed: {exc}"
            self._notify("execution.error", message, level="error")
            self.progress_message = "Execution failed"
            return

        self._notify("execution.complete", "Execution finished from TUI.")
        self._append_status("Execution complete")
        self.progress_message = "Execution complete"

    def _handle_mode(
        self, mode: str, overrides: Mapping[str, object], friendly: Optional[str] = None
    ) -> None:
        label = friendly or mode
        if not self._ensure_inventory():
            return

        payload: MutableMapping[str, object] = dict(overrides)
        payload.setdefault("mode", mode)

        self.progress_message = f"Planning {label}..."
        self._notify("plan.mode_start", f"Planning {label} run.", overrides=dict(payload))
        self._render()

        try:
            plan_data = self.planner(self.last_inventory or {}, payload)
        except Exception as exc:  # pragma: no cover - defensive logging
            message = f"{label.title()} planning failed: {exc}"
            self._notify("plan.mode_error", message, level="error")
            self.progress_message = f"{label.title()} failed"
            return

        self.last_plan = plan_data
        summary = plan_module.summarize_plan(plan_data)
        self._notify("plan.mode_ready", f"Plan ready for {label}.", summary=summary)
        for line in _format_plan(plan_data)[:6]:
            self._append_status(f"  {line}")

        if self.last_inventory is not None:
            payload["inventory"] = self.last_inventory

        self.progress_message = f"Executing {label}..."
        self._notify("execution.mode_start", f"Executing {label} run.", overrides=dict(payload))
        self._render()

        try:
            self.executor(plan_data, payload)
        except Exception as exc:  # pragma: no cover - defensive logging
            message = f"{label.title()} execution failed: {exc}"
            self._notify("execution.mode_error", message, level="error")
            self.progress_message = f"{label.title()} failed"
            return

        if payload.get("mode") == "diagnose":
            self._append_status("Diagnostics captured; no actions executed.")
            self.progress_message = "Diagnostics complete"
            self._notify("execution.diagnostics", "Diagnostics complete.")
        else:
            self._append_status("Execution complete")
            self.progress_message = "Execution complete"
            self._notify("execution.mode_complete", f"{label.title()} complete.")

    def _handle_targeted(self) -> None:
        if not self._ensure_inventory():
            return

        versions_raw = _read_input_line("Target versions (comma separated): ")
        targets = [item.strip() for item in versions_raw.split(",") if item.strip()]
        if not targets:
            self._notify(
                "targeted.cancel",
                "Targeted scrub aborted (no versions provided).",
                level="warning",
            )
            self.progress_message = "Targeted cancelled"
            return

        include_raw = _read_input_line(
            "Optional components (visio,project,onenote): "
        ).strip()

        overrides: MutableMapping[str, object] = {
            "mode": f"target:{','.join(targets)}",
            "target": ",".join(targets),
        }
        if include_raw:
            overrides["include"] = include_raw

        self._notify(
            "targeted.start",
            "Starting targeted scrub run.",
            targets=overrides.get("target"),
            include=overrides.get("include"),
        )
        self._handle_mode("targeted", overrides)

    def _handle_settings(self) -> None:
        args = self.app_state.get("args")
        details = [
            f"Dry-run: {bool(getattr(args, 'dry_run', False))}",
            f"Create restore point: {not bool(getattr(args, 'no_restore_point', False))}",
            f"License cleanup enabled: {not bool(getattr(args, 'no_license', False))}",
            f"Keep templates: {bool(getattr(args, 'keep_templates', False))}",
            f"Log directory: {getattr(args, 'logdir', '(default)')}",
            f"Backup directory: {getattr(args, 'backup', '(disabled)')}",
            "Timeout: "
            + (
                f"{getattr(args, 'timeout')} seconds"
                if getattr(args, "timeout", None) is not None
                else "(default)"
            ),
        ]
        for line in details:
            self._append_status(line)
        self._notify("settings.display", "Settings displayed in TUI.")
        self.progress_message = "Settings displayed"

    def _handle_logs(self) -> None:
        args = self.app_state.get("args")
        logdir = getattr(args, "logdir", None)
        message = f"Logs directory: {logdir or '(default)'}"
        self._append_status(message)
        self._notify("logs.info", message)
        self.progress_message = "Log details displayed"

    def _append_status(self, message: str) -> None:
        if self.status_lines and self.status_lines[-1] == message:
            return
        self.status_lines.append(message)
        limit = 24 if self.compact_layout else 32
        if len(self.status_lines) > limit:
            self.status_lines[:] = self.status_lines[-limit:]

    def _ensure_inventory(self) -> bool:
        if self.last_inventory is not None:
            return True
        self._handle_detect()
        return self.last_inventory is not None

    def _notify(
        self, event: str, message: str, *, level: str = "info", **payload: object
    ) -> None:
        human_logger = self.human_logger
        if human_logger is not None:
            log_func = getattr(human_logger, level, human_logger.info)
            log_func(message)

        machine_logger = self.machine_logger
        if machine_logger is not None:
            extra: dict[str, object] = {"event": "ui_progress", "name": event}
            if message:
                extra["message"] = message
            if payload:
                extra["data"] = dict(payload)
            machine_logger.info("ui_progress", extra=extra)

        record = {"event": event, "message": message}
        if payload:
            record["data"] = dict(payload)

        if callable(self.emit_event):
            try:
                self.emit_event(event, message=message, **payload)  # type: ignore[misc]
            except Exception:  # pragma: no cover - defensive path
                self.event_queue.append(record)
        else:
            self.event_queue.append(record)

        self._append_status(message)

    def _drain_events(self) -> bool:
        updated = False
        while self.event_queue:
            try:
                event = self.event_queue.popleft()
            except IndexError:
                break
            if not isinstance(event, Mapping):
                continue
            message = event.get("message")
            if message:
                self._append_status(str(message))
                updated = True
            progress = event.get("data")
            if isinstance(progress, Mapping) and progress.get("status"):
                self.progress_message = str(progress["status"])
        return updated


def run_tui(app_state: Mapping[str, object]) -> None:
    """!
    @brief Convenience wrapper to create and run the TUI controller.
    """

    OfficeJanitorTUI(app_state).run()


def _supports_ansi(stream: Optional[object] = None) -> bool:
    """!
    @brief Determine whether the current stdout supports ANSI escape sequences.
    """

    target = stream if stream is not None else sys.stdout
    if not hasattr(target, "isatty"):
        return False
    try:
        if not target.isatty():  # type: ignore[operator]
            return False
    except Exception:
        return False

    if os.name != "nt":
        return True
    return bool(
        os.environ.get("WT_SESSION")
        or os.environ.get("ANSICON")
        or os.environ.get("TERM_PROGRAM")
        or os.environ.get("ConEmuANSI")
    )


def _clear_screen() -> None:
    sys.stdout.write("\x1b[2J\x1b[H")
    sys.stdout.flush()


def _divider(width: int) -> str:
    return "-" * width


def _format_inventory(inventory: Mapping[str, object]) -> List[str]:
    lines: List[str] = []
    for key, items in inventory.items():
        try:
            count = len(items)  # type: ignore[arg-type]
        except TypeError:
            count = len(list(items))  # type: ignore[arg-type]
        lines.append(f"{key:<16} {count:>4}")
    if not lines:
        lines.append("No data")
    return lines


def _summarize_inventory(inventory: Mapping[str, object]) -> dict[str, int]:
    summary: dict[str, int] = {}
    for key, items in inventory.items():
        try:
            count = len(items)  # type: ignore[arg-type]
        except TypeError:
            count = len(list(items))  # type: ignore[arg-type]
        summary[str(key)] = count
    return summary


def _format_plan(plan_data: Optional[list[dict]]) -> List[str]:
    if not plan_data:
        return ["Plan not created"]

    summary = plan_module.summarize_plan(plan_data)
    lines: List[str] = [
        f"Steps: {summary.get('total_steps', 0)} (actionable {summary.get('actionable_steps', 0)})",
    ]

    mode = summary.get("mode")
    dry_run = summary.get("dry_run")
    if mode:
        lines.append(f"Mode: {mode}{' [dry-run]' if dry_run else ''}")

    target_versions = summary.get("target_versions") or []
    if target_versions:
        lines.append("Targets: " + ", ".join(str(item) for item in target_versions))

    discovered = summary.get("discovered_versions") or []
    if discovered:
        lines.append("Detected: " + ", ".join(str(item) for item in discovered))

    uninstall_versions = summary.get("uninstall_versions") or []
    if uninstall_versions:
        lines.append("Uninstalls: " + ", ".join(str(item) for item in uninstall_versions))

    cleanup_categories = summary.get("cleanup_categories") or []
    if cleanup_categories:
        lines.append("Cleanup: " + ", ".join(str(item) for item in cleanup_categories))

    requested_components = summary.get("requested_components") or []
    if requested_components:
        lines.append("Include: " + ", ".join(str(item) for item in requested_components))

    unsupported_components = summary.get("unsupported_components") or []
    if unsupported_components:
        lines.append(
            "Unsupported include: " + ", ".join(str(item) for item in unsupported_components)
        )

    categories = summary.get("categories") or {}
    if categories:
        formatted = ", ".join(f"{key}={value}" for key, value in categories.items())
        lines.append("Categories: " + formatted)

    return lines


def _read_input_line(prompt: str) -> str:
    return input(prompt)


def _spinner(duration: float, message: str) -> None:
    frames = ["-", "\\", "|", "/"]
    start = time.monotonic()
    index = 0
    while time.monotonic() - start < duration:
        sys.stdout.write(f"\r{message} {frames[index % len(frames)]}")
        sys.stdout.flush()
        time.sleep(0.1)
        index += 1
    sys.stdout.write("\r" + " " * (len(message) + 2) + "\r")


def _default_key_reader() -> str:
    try:  # pragma: no cover - depends on Windows availability
        import msvcrt

        if msvcrt.kbhit():
            char = msvcrt.getwch()
        else:
            char = msvcrt.getwch()
        return char
    except Exception:
        return _read_input_line("Command (d=detect, p=plan, r=run, l=logs, q=quit): ").strip()[:1]


