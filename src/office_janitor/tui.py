"""!
@file tui.py
@brief Text-based user interface (TUI) engine.

@details Implements the ANSI/VT driven interface described in the project
specification. The implementation keeps dependencies to the standard library
only while providing a co-operative event loop that drains orchestrator
progress events and handles keyboard commands. The layout follows a header,
navigation column, and tabbed content panes so that additional widgets can be
rendered predictably across platforms.

This module contains the main OfficeJanitorTUI class. Rendering is provided
by TUIRendererMixin from tui_render.py, action handlers from TUIActionsMixin
in tui_actions.py, and helper functions are in tui_helpers.py.
"""

from __future__ import annotations

import time
from collections import deque
from collections.abc import Mapping, MutableMapping
from dataclasses import dataclass, field
from typing import Any, Callable

from . import constants
from .tui_actions import TUIActionsMixin
from .tui_helpers import (
    decode_key,
    default_key_reader,
    format_inventory,
    read_input_line,
    supports_ansi,
)
from .tui_render import TUIRendererMixin

# Re-export for backward compatibility
from .tui_helpers import (
    decode_key as _decode_key,
    default_key_reader as _default_key_reader,
    divider as _divider,
    format_inventory as _format_inventory,
    format_plan as _format_plan,
    read_input_line as _read_input_line,
    render_progress_bar,
    spinner as _spinner,
    strip_ansi as _strip_ansi,
    summarize_inventory as _summarize_inventory,
    supports_ansi as _supports_ansi,
    clear_screen as _clear_screen,
)

try:  # pragma: no cover - Windows specific
    import msvcrt as _msvcrt
except ImportError:  # pragma: no cover - non-Windows hosts
    _msvcrt = None

msvcrt: Any = _msvcrt


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------


@dataclass
class NavigationItem:
    """!
    @brief Describes an entry in the navigation column.
    """

    name: str
    label: str
    action: Callable[[], None] | None = None
    quit_on_activate: bool = False


@dataclass
class PaneContext:
    """!
    @brief Tracks cursor position and data for a pane.
    """

    name: str
    cursor: int = 0
    lines: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Main TUI class
# ---------------------------------------------------------------------------


class OfficeJanitorTUI(TUIRendererMixin, TUIActionsMixin):
    """!
    @brief Controller coordinating rendering, events, and orchestrator calls.
    """

    def __init__(self, app_state: Mapping[str, object]) -> None:
        self.app_state: MutableMapping[str, object] = dict(app_state)
        self.human_logger = self.app_state.get("human_logger")
        self.machine_logger = self.app_state.get("machine_logger")
        self.detector: Callable[[], Mapping[str, object]] = self.app_state["detector"]
        self.planner: Callable[
            [Mapping[str, object], Mapping[str, object] | None], list[dict[str, object]]
        ] = self.app_state["planner"]
        self.executor: Callable[
            [list[dict[str, object]], Mapping[str, object] | None], bool | None
        ] = self.app_state["executor"]
        confirm_callable = self.app_state.get("confirm")
        self._confirm_requestor: Callable[..., bool] | None = (
            confirm_callable if callable(confirm_callable) else None
        )
        queue_obj = self.app_state.get("event_queue")
        if isinstance(queue_obj, deque):
            self.event_queue: deque[dict[str, object]] = queue_obj
        else:
            self.event_queue = deque()
            self.app_state["event_queue"] = self.event_queue
        self.emit_event = self.app_state.get("emit_event")
        self.last_inventory: Mapping[str, object] | None = None
        self.last_plan: list[dict[str, object]] | None = None
        self.status_lines: list[str] = []
        self.progress_message = "Ready"
        self.log_lines: list[str] = []
        self.ansi_supported = supports_ansi() and not bool(
            getattr(self.app_state.get("args"), "no_color", False)
        )
        self._key_reader: Callable[[], str] | None = self.app_state.get("key_reader")
        args = self.app_state.get("args")
        refresh_ms = getattr(args, "tui_refresh", 120) if args is not None else 120
        try:
            refresh_value = float(refresh_ms) / 1000.0
        except Exception:
            refresh_value = 0.12
        self.refresh_interval = 0.05 if refresh_value <= 0 else refresh_value
        self.compact_layout = (
            bool(getattr(args, "tui_compact", False)) if args is not None else False
        )
        self._running = True
        self.navigation: list[NavigationItem] = [
            NavigationItem("detect", "Detect inventory", action=self._handle_detect),
            NavigationItem("auto", "Auto scrub everything", action=self._handle_auto_all),
            NavigationItem("targeted", "Targeted scrub", action=self._prepare_targeted),
            NavigationItem("cleanup", "Cleanup only", action=self._handle_cleanup_only),
            NavigationItem("diagnostics", "Diagnostics only", action=self._handle_diagnostics),
            NavigationItem("plan", "Build plan", action=None),
            NavigationItem("run", "Run plan", action=self._handle_run),
            NavigationItem("logs", "Live logs", action=self._handle_logs),
            NavigationItem("settings", "Settings", action=None),
            NavigationItem("quit", "Quit", quit_on_activate=True),
        ]
        self.focus_area = "nav"
        self.nav_index = 0
        self.active_tab = self.navigation[0].name
        self.panes: dict[str, PaneContext] = {
            item.name: PaneContext(item.name) for item in self.navigation
        }
        self.plan_overrides: dict[str, bool] = {
            "include_visio": False,
            "include_project": False,
            "include_onenote": False,
        }
        self.target_overrides: dict[str, bool] = {
            str(version): False for version in constants.SUPPORTED_TARGETS
        }
        self.list_filters: dict[str, str] = {}
        args_defaults = {
            "dry_run": bool(getattr(args, "dry_run", False)),
            "create_restore_point": not bool(getattr(args, "no_restore_point", False)),
            "license_cleanup": not bool(getattr(args, "no_license", False)),
            "keep_templates": bool(getattr(args, "keep_templates", False)),
        }
        self.settings_overrides: dict[str, bool] = dict(args_defaults)
        self.last_overrides: dict[str, object] | None = None

    # -----------------------------------------------------------------------
    # Confirmation input (overrides stub in TUIActionsMixin)
    # -----------------------------------------------------------------------

    def _prompt_confirmation_input(self, label: str, prompt: str) -> str:
        """!
        @brief Display a confirmation prompt and collect a response inside the TUI.
        """

        prompt_text = prompt.strip()
        if prompt_text:
            self._append_status(prompt_text)

        previous_message = self.progress_message
        display_prompt = prompt_text or f"Confirm {label}"
        self.progress_message = display_prompt
        self._render()

        while self._running:
            if self._drain_events():
                self._render()
                continue

            command = self._read_command()
            if not command:
                time.sleep(self.refresh_interval)
                continue

            normalized = command.lower()
            if command == "enter":
                self.progress_message = previous_message
                self._render()
                return ""

            if len(command) == 1 and normalized in {"y", "n"}:
                self.progress_message = previous_message
                self._render()
                return normalized

            if normalized in {"quit", "escape"}:
                self.progress_message = previous_message
                self._render()
                return "n"

            self._append_status("Press Y to confirm or N to cancel.")
            self._render()

        self.progress_message = previous_message
        self._render()
        return "n"

    # -----------------------------------------------------------------------
    # Main event loop
    # -----------------------------------------------------------------------

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

        if not self.ansi_supported and not supports_ansi():
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

            self._handle_key(command)
            self._render()

    # -----------------------------------------------------------------------
    # Key handling
    # -----------------------------------------------------------------------

    def _handle_key(self, command: str) -> None:
        """!
        @brief Interpret a normalized key command and update state.
        """

        if command in {"quit", "escape"}:
            self.progress_message = "Exiting..."
            self._notify("tui.exit", "User requested exit from TUI.")
            self._running = False
            return

        if command == "tab":
            self.focus_area = "content" if self.focus_area == "nav" else "nav"
            return

        if self.focus_area == "nav":
            if command == "down":
                self._move_nav(1)
                return
            if command == "up":
                self._move_nav(-1)
                return
            if command in {"enter", "space"}:
                self._activate_nav()
                return
            if command == "f10":
                self._handle_run()
                return
            if command == "f1":
                self._show_help()
                return

        self._handle_content_key(command)

    def _handle_content_key(self, command: str) -> None:
        """!
        @brief Handle key presses that interact with the active pane.
        """

        pane = self.panes.get(self.active_tab)
        if pane is None:
            return

        self._ensure_pane_lines(pane)

        if command == "up":
            pane.cursor = max(0, pane.cursor - 1)
            return
        if command == "down":
            pane.cursor = min(max(len(pane.lines) - 1, 0), pane.cursor + 1)
            return
        if command == "page_down":
            pane.cursor = min(max(len(pane.lines) - 1, 0), pane.cursor + 5)
            return
        if command == "page_up":
            pane.cursor = max(0, pane.cursor - 5)
            return
        if command == "f10":
            if self.active_tab == "targeted":
                self._handle_targeted(execute=True)
            elif self.active_tab == "auto":
                self._handle_auto_all()
            elif self.active_tab == "cleanup":
                self._handle_cleanup_only()
            elif self.active_tab == "diagnostics":
                self._handle_diagnostics()
            else:
                self._handle_run()
            return
        if command == "f1":
            self._show_help()
            return
        if command == "enter":
            if self.active_tab == "plan":
                self._handle_plan()
            elif self.active_tab == "run":
                self._handle_run()
            elif self.active_tab == "logs":
                self._append_status("Logs reviewed")
            elif self.active_tab == "settings":
                self._append_status("Settings confirmed")
            elif self.active_tab == "targeted":
                self._handle_targeted(execute=True)
            elif self.active_tab == "auto":
                self._handle_auto_all()
            elif self.active_tab == "cleanup":
                self._handle_cleanup_only()
            elif self.active_tab == "diagnostics":
                self._handle_diagnostics()
            return
        if command == "space":
            if self.active_tab == "plan":
                self._toggle_plan_option(pane.cursor)
            elif self.active_tab == "settings":
                self._toggle_setting(pane.cursor)
            elif self.active_tab == "targeted":
                self._toggle_target(pane.cursor)
            return
        if command == "/":
            self._prompt_filter(pane)
            self._ensure_pane_lines(pane)
            return

    # -----------------------------------------------------------------------
    # Pane state management
    # -----------------------------------------------------------------------

    def _ensure_pane_lines(self, pane: PaneContext) -> list[tuple[str, str]]:
        """Populate pane lines based on current state."""
        entries: list[tuple[str, str]] = []
        if pane.name == "plan":
            entries = [
                (
                    key,
                    f"{'[x]' if enabled else '[ ]'} {key.replace('_', ' ').title()}",
                )
                for key, enabled in self.plan_overrides.items()
            ]
        elif pane.name == "targeted":
            entries = [
                (
                    version,
                    f"{'[x]' if enabled else '[ ]'} Office {version}",
                )
                for version, enabled in self.target_overrides.items()
            ]
        elif pane.name == "settings":
            entries = [
                (
                    key,
                    f"{'[x]' if enabled else '[ ]'} {key.replace('_', ' ').title()}",
                )
                for key, enabled in self.settings_overrides.items()
            ]
        elif pane.name == "logs":
            entries = [(line, line) for line in self.log_lines]
        elif pane.name == "run":
            recent = self.status_lines[-10:]
            entries = [(line, line) for line in recent]
        elif pane.name == "detect" and self.last_inventory is not None:
            entries = [(line, line) for line in format_inventory(self.last_inventory)]
        elif pane.name in {"auto", "cleanup", "diagnostics"}:
            entries = []
        else:
            entries = [(line, line) for line in pane.lines]

        return self._filter_entries(pane, entries)

    def _get_pane_filter(self, pane_name: str) -> str:
        """Get the current filter for a pane."""
        return self.list_filters.get(pane_name, "")

    def _set_pane_filter(self, pane_name: str, value: str) -> None:
        """Set or clear the filter for a pane."""
        normalized = value.strip()
        if normalized:
            self.list_filters[pane_name] = normalized
        else:
            self.list_filters.pop(pane_name, None)

    def _filter_entries(
        self, pane: PaneContext, entries: list[tuple[str, str]]
    ) -> list[tuple[str, str]]:
        """Apply filtering to pane entries."""
        filter_text = self._get_pane_filter(pane.name)
        if filter_text:
            lowered = filter_text.lower()
            filtered = [entry for entry in entries if lowered in entry[1].lower()]
        else:
            filtered = list(entries)

        pane.lines = [entry[0] for entry in filtered]
        if pane.lines:
            pane.cursor = max(0, min(pane.cursor, len(pane.lines) - 1))
        else:
            pane.cursor = 0
        return filtered

    # -----------------------------------------------------------------------
    # Navigation
    # -----------------------------------------------------------------------

    def _move_nav(self, offset: int) -> None:
        """Move navigation cursor by offset."""
        size = len(self.navigation)
        self.nav_index = (self.nav_index + offset) % size
        self.active_tab = self.navigation[self.nav_index].name
        self.focus_area = "nav"

    def _activate_nav(self) -> None:
        """Activate the currently selected navigation item."""
        item = self.navigation[self.nav_index]
        if item.quit_on_activate:
            self._running = False
            self.progress_message = "Exiting..."
            self._notify("tui.exit", "User requested exit from TUI.")
            return
        self.active_tab = item.name
        if item.action is not None:
            item.action()
        self.focus_area = "content"

    def _read_command(self) -> str:
        """Read and decode a key command."""
        reader = self._key_reader or default_key_reader
        try:
            raw = reader()
        except StopIteration:
            self._running = False
            return "quit"
        except Exception:
            return ""
        return decode_key(raw)

    # -----------------------------------------------------------------------
    # Status and event handling
    # -----------------------------------------------------------------------

    def _append_status(self, message: str) -> None:
        """Append a message to the status log."""
        if self.status_lines and self.status_lines[-1] == message:
            return
        self.status_lines.append(message)
        limit = 24 if self.compact_layout else 32
        if len(self.status_lines) > limit:
            self.status_lines[:] = self.status_lines[-limit:]

    def _notify(self, event: str, message: str, *, level: str = "info", **payload: object) -> None:
        """Send a notification through loggers and event queue."""
        human_logger = self.human_logger
        if human_logger is not None:
            log_func = getattr(human_logger, level, human_logger.info)
            log_func(message)

        machine_logger = self.machine_logger
        if machine_logger is not None:
            extra: dict[str, object] = {"event": "ui_progress", "event_name": event}
            if message:
                extra["log_message"] = message
            if payload:
                extra["data"] = dict(payload)
            machine_logger.info("ui_progress", extra=extra)

        record = {"event": event, "message": message}
        if payload:
            record["data"] = dict(payload)

        if callable(self.emit_event):
            try:
                self.emit_event(event, message=message, **payload)
            except Exception:  # pragma: no cover - defensive path
                self.event_queue.append(record)
        else:
            self.event_queue.append(record)

        self._append_status(message)

    def _drain_events(self) -> bool:
        """Process pending events from the event queue."""
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
            data = event.get("data")
            if isinstance(data, Mapping):
                status = data.get("status")
                if status:
                    self.progress_message = str(status)
                inventory = data.get("inventory")
                if isinstance(inventory, Mapping):
                    self.last_inventory = inventory
                plan_data = data.get("plan")
                if isinstance(plan_data, list):
                    self.last_plan = list(plan_data)
                log_line = data.get("log_line")
                if log_line:
                    self._append_log(str(log_line))
                toggle_updates = data.get("settings")
                if isinstance(toggle_updates, Mapping):
                    for key, value in toggle_updates.items():
                        if key in self.settings_overrides:
                            self.settings_overrides[key] = bool(value)
            event_name = event.get("event")
            if event_name == "log.line" and "message" in event:
                self._append_log(str(event["message"]))
        return updated

    def _append_log(self, line: str) -> None:
        """Append a line to the log buffer."""
        self.log_lines.append(line)
        limit = 200
        if len(self.log_lines) > limit:
            self.log_lines[:] = self.log_lines[-limit:]


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def run_tui(app_state: Mapping[str, object]) -> None:
    """!
    @brief Convenience wrapper to create and run the TUI controller.
    """

    OfficeJanitorTUI(app_state).run()
