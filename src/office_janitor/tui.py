"""!
@brief Text-based user interface (TUI) engine.
@details Implements the ANSI/VT driven interface with panes, widgets, and event
handling outlined in the specification for rich interactive sessions.
"""
from __future__ import annotations

import os
import sys
import time
from typing import Callable, List, Mapping, MutableMapping, Optional

from . import plan as plan_module


class OfficeJanitorTUI:
    """!
    @brief Placeholder TUI controller coordinating rendering and event handling.
    """

    def __init__(self, app_state: Mapping[str, object]) -> None:
        self.app_state: MutableMapping[str, object] = dict(app_state)
        self.human_logger = self.app_state.get("human_logger")
        self.machine_logger = self.app_state.get("machine_logger")
        self.detector: Callable[[], Mapping[str, object]] = self.app_state["detector"]  # type: ignore[assignment]
        self.planner: Callable[[Mapping[str, object], Mapping[str, object] | None], list[dict]] = self.app_state["planner"]  # type: ignore[assignment]
        self.executor: Callable[[list[dict], Mapping[str, object] | None], None] = self.app_state["executor"]  # type: ignore[assignment]
        self.last_inventory: Optional[Mapping[str, object]] = None
        self.last_plan: Optional[list[dict]] = None
        self.status_lines: List[str] = []
        self.progress_message = "Ready"
        self._key_reader: Optional[Callable[[], str]] = self.app_state.get("key_reader")  # type: ignore[assignment]

    def run(self) -> None:
        """!
        @brief Enter the TUI event loop.
        """

        args = self.app_state.get("args")
        if getattr(args, "quiet", False) or getattr(args, "json", False):
            if self.human_logger:
                self.human_logger.info(
                    "Interactive TUI suppressed because quiet/json output mode was requested."
                )
            return

        if not _supports_ansi() or getattr(args, "no_color", False):
            from . import ui

            if self.human_logger:
                self.human_logger.info("Falling back to plain CLI menu (ANSI unavailable).")
            ui.run_cli(self.app_state)
            return

        self._render()
        while True:
            command = self._read_command()
            if not command:
                continue
            if command in {"q", "Q", "\u001b"}:
                self.progress_message = "Exiting..."
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
                self._handle_mode("auto-all", {"mode": "auto-all", "auto_all": True}, friendly="auto scrub")
            elif command in {"t", "T"}:
                self._handle_targeted()
            elif command in {"c", "C"}:
                self._handle_mode("cleanup-only", {"mode": "cleanup-only", "cleanup_only": True}, friendly="cleanup")
            elif command in {"g", "G"}:
                self._handle_mode("diagnose", {"mode": "diagnose", "diagnose": True}, friendly="diagnostics")
            elif command in {"s", "S"}:
                self._handle_settings()
            else:
                self.progress_message = f"Unknown command: {command!r}"
            self._render()

    def _render(self) -> None:
        width = 90
        left_width = 32
        _clear_screen()
        header = f"Office Janitor — {self.progress_message}"
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
        ] + self.status_lines[-10:]

        inventory_lines = _format_inventory(self.last_inventory) if self.last_inventory else ["No inventory"]
        plan_lines = _format_plan(self.last_plan)
        right_lines = ["Inventory summary:"] + inventory_lines + ["", "Plan summary:"] + plan_lines

        max_lines = max(len(left_lines), len(right_lines))
        for index in range(max_lines):
            left_text = left_lines[index] if index < len(left_lines) else ""
            right_text = right_lines[index] if index < len(right_lines) else ""
            sys.stdout.write(f"{left_text.ljust(left_width)} {right_text}\n")

        sys.stdout.write(_divider(width) + "\n")
        sys.stdout.write(
            "Commands: d=Detect p=Plan r=Run a=Auto t=Target c=Cleanup g=Diagnose l=Logs s=Settings q=Quit\n"
        )
        sys.stdout.flush()

    def _read_command(self) -> str:
        reader = self._key_reader or _default_key_reader
        try:
            command = reader()
        except Exception:
            return ""
        return command.strip()[:1] if isinstance(command, str) else ""

    def _handle_detect(self) -> None:
        self.progress_message = "Detecting inventory..."
        self._render()
        try:
            inventory = self.detector()
        except Exception as exc:  # pragma: no cover - defensive logging
            message = f"Detection failed: {exc}"
            self._append_status(message)
            if self.human_logger:
                self.human_logger.error(message)
            self.progress_message = "Detection failed"
            return

        self.last_inventory = inventory
        summary = _format_inventory(inventory)
        self._append_status("Inventory updated")
        for line in summary:
            self._append_status(f"  {line}")
        self.progress_message = "Inventory ready"

    def _handle_plan(self) -> None:
        if self.last_inventory is None:
            self._handle_detect()
            if self.last_inventory is None:
                return

        self.progress_message = "Planning actions..."
        self._render()
        try:
            plan_data = self.planner(self.last_inventory, None)
        except Exception as exc:  # pragma: no cover - defensive logging
            message = f"Plan failed: {exc}"
            self._append_status(message)
            if self.human_logger:
                self.human_logger.error(message)
            self.progress_message = "Plan failed"
            return

        self.last_plan = plan_data
        self._append_status(f"Plan ready with {len(plan_data)} steps")
        for line in _format_plan(plan_data)[:6]:
            self._append_status(f"  {line}")
        self.progress_message = "Plan ready"

    def _handle_run(self) -> None:
        if self.last_plan is None:
            self._handle_plan()
            if self.last_plan is None:
                return

        self.progress_message = "Executing plan..."
        self._render()
        try:
            _spinner(0.2, "Preparing")
            self.executor(self.last_plan, None)
        except Exception as exc:  # pragma: no cover - defensive logging
            message = f"Execution failed: {exc}"
            self._append_status(message)
            if self.human_logger:
                self.human_logger.error(message)
            self.progress_message = "Execution failed"
            return

        self._append_status("Execution complete")
        self.progress_message = "Execution complete"

    def _handle_mode(
        self, mode: str, overrides: Mapping[str, object], friendly: Optional[str] = None
    ) -> None:
        label = friendly or mode
        if not self._ensure_inventory():
            return

        payload: MutableMapping[str, object] = dict(overrides)
        if "mode" not in payload:
            payload["mode"] = mode

        self.progress_message = f"Preparing {label}..."
        self._render()

        try:
            plan_data = self.planner(self.last_inventory or {}, payload)
        except Exception as exc:  # pragma: no cover - defensive logging
            message = f"{label.title()} planning failed: {exc}"
            self._append_status(message)
            if self.human_logger:
                self.human_logger.error(message)
            self.progress_message = f"{label.title()} failed"
            return

        self.last_plan = plan_data
        self._append_status(f"Plan ready with {len(plan_data)} steps for {label}")
        for line in _format_plan(plan_data)[:6]:
            self._append_status(f"  {line}")

        if self.last_inventory is not None:
            payload["inventory"] = self.last_inventory

        self.progress_message = f"Executing {label}..."
        self._render()

        try:
            self.executor(plan_data, payload)
        except Exception as exc:  # pragma: no cover - defensive logging
            message = f"{label.title()} execution failed: {exc}"
            self._append_status(message)
            if self.human_logger:
                self.human_logger.error(message)
            self.progress_message = f"{label.title()} failed"
            return

        if payload.get("mode") == "diagnose":
            self._append_status("Diagnostics captured; no actions executed.")
            self.progress_message = "Diagnostics complete"
        else:
            self._append_status("Execution complete")
            self.progress_message = "Execution complete"

    def _handle_targeted(self) -> None:
        if not self._ensure_inventory():
            return

        versions_raw = _read_input_line("Target versions (comma separated): ")
        targets = [item.strip() for item in versions_raw.split(",") if item.strip()]
        if not targets:
            self._append_status("No target versions entered; aborting targeted run.")
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

        self._handle_mode("targeted", overrides)

    def _handle_settings(self) -> None:
        args = self.app_state.get("args")
        details = [
            f"Dry-run: {bool(getattr(args, 'dry_run', False))}",
            f"Create restore point: {not bool(getattr(args, 'no_restore_point', False))}",
            f"Log directory: {getattr(args, 'logdir', '(default)')}",
            f"Backup directory: {getattr(args, 'backup', '(disabled)')}",
        ]
        for line in details:
            self._append_status(line)
        self.progress_message = "Settings displayed"

    def _handle_logs(self) -> None:
        args = self.app_state.get("args")
        logdir = getattr(args, "logdir", None)
        message = f"Logs directory: {logdir or '(default)'}"
        self._append_status(message)
        self.progress_message = "Log details displayed"

    def _append_status(self, message: str) -> None:
        self.status_lines.append(message)
        if len(self.status_lines) > 20:
            self.status_lines[:] = self.status_lines[-20:]

    def _ensure_inventory(self) -> bool:
        if self.last_inventory is not None:
            return True
        self._handle_detect()
        return self.last_inventory is not None


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


def _move_cursor(row: int, col: int) -> None:
    sys.stdout.write(f"\x1b[{row};{col}H")


def _divider(width: int) -> str:
    return "-" * width


def _format_inventory(inventory: Mapping[str, object]) -> List[str]:
    lines: List[str] = []
    for key, items in inventory.items():
        try:
            count = len(items)  # type: ignore[arg-type]
        except TypeError:
            count = len(list(items))  # type: ignore[arg-type]
        lines.append(f"{key:<12} {count:>5}")
    if not lines:
        lines.append("No data")
    return lines


def _format_plan(plan_data: Optional[list[dict]]) -> List[str]:
    if not plan_data:
        return ["Plan not created"]

    summary = plan_module.summarize_plan(plan_data)
    lines: List[str] = ["Plan summary:"]
    lines.append(
        " Steps: {total} (actionable {actionable})".format(
            total=summary.get("total_steps", 0),
            actionable=summary.get("actionable_steps", 0),
        )
    )

    mode = summary.get("mode")
    dry_run = summary.get("dry_run")
    if mode:
        lines.append(f" Mode: {mode}{' [dry-run]' if dry_run else ''}")

    target_versions = summary.get("target_versions") or []
    discovered = summary.get("discovered_versions") or []
    if target_versions:
        lines.append(" Targets: " + ", ".join(str(item) for item in target_versions))
    if discovered:
        lines.append(" Detected: " + ", ".join(str(item) for item in discovered))

    uninstall_versions = summary.get("uninstall_versions") or []
    if uninstall_versions:
        lines.append(" Uninstalls: " + ", ".join(str(item) for item in uninstall_versions))

    cleanup_categories = summary.get("cleanup_categories") or []
    if cleanup_categories:
        lines.append(" Cleanup: " + ", ".join(str(item) for item in cleanup_categories))

    requested_components = summary.get("requested_components") or []
    if requested_components:
        lines.append(" Include: " + ", ".join(str(item) for item in requested_components))
    unsupported_components = summary.get("unsupported_components") or []
    if unsupported_components:
        lines.append(
            " Unsupported include: " + ", ".join(str(item) for item in unsupported_components)
        )

    categories = summary.get("categories") or {}
    if categories:
        formatted = ", ".join(f"{key}={value}" for key, value in categories.items())
        lines.append(" Categories: " + formatted)

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


