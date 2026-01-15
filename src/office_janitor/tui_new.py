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
by TUIRendererMixin from tui_render.py, and helper functions are in tui_helpers.py.
"""

from __future__ import annotations

import time
from collections import deque
from collections.abc import Mapping, MutableMapping
from dataclasses import dataclass, field
from typing import Any, Callable

from . import constants
from . import plan as plan_module
from .tui_helpers import (
    decode_key,
    default_key_reader,
    format_inventory,
    format_plan,
    read_input_line,
    spinner,
    summarize_inventory,
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


class OfficeJanitorTUI(TUIRendererMixin):
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
    # Confirmation and input handling
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

    def _confirm_execution(self, label: str, overrides: MutableMapping[str, object]) -> bool:
        """!
        @brief Coordinate scrub confirmation for interactive executions.
        """

        request_confirmation = self._confirm_requestor
        if request_confirmation is None:
            overrides["confirmed"] = True
            return True

        args = self.app_state.get("args")
        dry_run = bool(getattr(args, "dry_run", False))
        if "dry_run" in overrides:
            dry_run = bool(overrides["dry_run"])

        force_override = bool(getattr(args, "force", False))
        if "force" in overrides:
            force_override = bool(overrides["force"])

        interactive_override = overrides.get("interactive")

        proceed = request_confirmation(
            dry_run=dry_run,
            force=force_override,
            input_func=lambda prompt: self._prompt_confirmation_input(label, prompt),
            interactive=(True if interactive_override is None else bool(interactive_override)),
        )

        if not proceed:
            message = f"{label.title()} cancelled"
            self._notify(
                "execution.cancelled",
                message,
                level="warning",
                reason="user_declined",
            )
            self.progress_message = message
            return False

        overrides["confirmed"] = True
        return True

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

    def _prompt_filter(self, pane: PaneContext) -> None:
        """Prompt user for filter text."""
        label = pane.name.replace("_", " ").title()
        previous_filter = self._get_pane_filter(pane.name)
        previous_message = self.progress_message
        response = ""
        try:
            self.progress_message = f"Filter {label}"
            self._render()
            try:
                response = read_input_line(f"Filter {label} (empty to clear): ")
            except (EOFError, KeyboardInterrupt):
                response = ""
            except Exception:
                response = ""
        finally:
            self.progress_message = previous_message

        new_value = response.strip()
        self._set_pane_filter(pane.name, new_value)
        applied_filter = self._get_pane_filter(pane.name)
        if applied_filter and applied_filter != previous_filter:
            self._append_status(f"{label} filter set to '{applied_filter}'")
        elif not applied_filter and previous_filter:
            self._append_status(f"{label} filter cleared")

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
    # Toggle handlers
    # -----------------------------------------------------------------------

    def _toggle_plan_option(self, index: int) -> None:
        """Toggle a plan option at the given index."""
        pane = self.panes.get("plan")
        if pane is None:
            return
        # Refresh pane.lines to get current filtered keys
        self._ensure_pane_lines(pane)
        keys = list(pane.lines) if pane.lines else []
        if not keys:
            self._append_status("No plan options available to toggle.")
            return
        safe_index = max(0, min(index, len(keys) - 1))
        key = keys[safe_index]
        if key not in self.plan_overrides:
            self._append_status(f"Unknown plan option: {key}")
            return
        self.plan_overrides[key] = not self.plan_overrides[key]
        self._append_status(f"Plan option '{key}' set to {self.plan_overrides[key]}")

    def _toggle_target(self, index: int) -> None:
        """Toggle a target version at the given index."""
        pane = self.panes.get("targeted")
        if pane is None:
            return
        # Refresh pane.lines to get current filtered keys
        self._ensure_pane_lines(pane)
        keys = list(pane.lines) if pane.lines else []
        if not keys:
            self._append_status("No target versions available to toggle.")
            return
        safe_index = max(0, min(index, len(keys) - 1))
        key = keys[safe_index]
        if key not in self.target_overrides:
            self._append_status(f"Unknown target version: {key}")
            return
        self.target_overrides[key] = not self.target_overrides[key]
        state = "selected" if self.target_overrides[key] else "cleared"
        self._append_status(f"Target version {key} {state}.")

    def _toggle_setting(self, index: int) -> None:
        """Toggle a setting at the given index."""
        pane = self.panes.get("settings")
        if pane is None:
            return
        # Refresh pane.lines to get current filtered keys
        self._ensure_pane_lines(pane)
        keys = list(pane.lines) if pane.lines else []
        if not keys:
            self._append_status("No settings available to toggle.")
            return
        safe_index = max(0, min(index, len(keys) - 1))
        key = keys[safe_index]
        if key not in self.settings_overrides:
            self._append_status(f"Unknown setting: {key}")
            return
        self.settings_overrides[key] = not self.settings_overrides[key]
        self._append_status(f"Setting '{key}' toggled to {self.settings_overrides[key]}")

    # -----------------------------------------------------------------------
    # Override collection
    # -----------------------------------------------------------------------

    def _collect_settings_overrides(self) -> dict[str, object]:
        """Collect settings overrides for plan execution."""
        overrides: dict[str, object] = {}
        toggles = self.settings_overrides
        overrides["dry_run"] = bool(toggles.get("dry_run", False))
        overrides["create_restore_point"] = bool(toggles.get("create_restore_point", False))
        overrides["keep_templates"] = bool(toggles.get("keep_templates", False))
        license_cleanup = bool(toggles.get("license_cleanup", True))
        overrides["license_cleanup"] = license_cleanup
        overrides["no_license"] = not license_cleanup
        return overrides

    def _collect_plan_overrides(self) -> dict[str, object]:
        """Collect plan overrides for plan building."""
        overrides: dict[str, object] = {}
        include_components: list[str] = []
        for key, enabled in self.plan_overrides.items():
            if not enabled:
                continue
            overrides[key] = True
            if key.startswith("include_"):
                include_components.append(key.split("_", 1)[1])
        if include_components:
            overrides["include"] = ",".join(include_components)
        return overrides

    def _combine_overrides(self, extra: Mapping[str, object] | None = None) -> dict[str, object]:
        """Combine all overrides into a single dictionary."""
        combined: dict[str, object] = {}
        combined.update(self._collect_settings_overrides())
        combined.update(self._collect_plan_overrides())
        if extra:
            combined.update({key: extra[key] for key in extra})
        return combined

    def _selected_targets(self) -> list[str]:
        """Get list of selected target versions."""
        return [version for version, enabled in self.target_overrides.items() if enabled]

    # -----------------------------------------------------------------------
    # Action handlers
    # -----------------------------------------------------------------------

    def _ensure_inventory(self) -> bool:
        """Ensure inventory is populated, running detection if needed."""
        if self.last_inventory is not None:
            return True
        self._handle_detect()
        return self.last_inventory is not None

    def _prepare_targeted(self) -> None:
        """Prepare for targeted scrub."""
        self._ensure_inventory()
        self.progress_message = "Configure targeted scrub"
        self._append_status("Targeted scrub ready: toggle versions with Space, F10 to run.")

    def _plan_and_optionally_execute(
        self,
        extra_overrides: Mapping[str, object] | None,
        *,
        label: str,
        execute: bool,
    ) -> None:
        """Build a plan and optionally execute it."""
        if not self._ensure_inventory():
            return

        combined = self._combine_overrides(extra_overrides)
        if "mode" not in combined:
            combined["mode"] = "interactive"

        self.progress_message = f"Planning {label}..."
        self._notify(
            "plan.start",
            f"Building plan for {label}.",
            overrides=dict(combined),
        )
        self._render()

        try:
            plan_data = self.planner(self.last_inventory, combined)
        except Exception as exc:  # pragma: no cover - defensive logging
            message = f"Plan failed: {exc}"
            self._notify("plan.error", message, level="error")
            self.progress_message = f"{label.title()} plan failed"
            return

        self.last_plan = plan_data
        self.last_overrides = dict(combined)
        summary = plan_module.summarize_plan(plan_data)
        self._notify(
            "plan.complete",
            f"{label.title()} plan ready.",
            summary=summary,
        )
        for line in format_plan(plan_data)[:6]:
            self._append_status(f"  {line}")

        if not execute:
            self.progress_message = f"{label.title()} plan ready"
            return

        payload = dict(combined)
        if self.last_inventory is not None:
            payload.setdefault("inventory", self.last_inventory)

        if not self._confirm_execution(label, payload):
            return

        self.progress_message = f"Executing {label}..."
        self._notify("execution.start", f"Executing {label} run.")
        self._render()

        try:
            spinner(0.2, "Preparing")
            execution_result = self.executor(plan_data, payload)
        except Exception as exc:  # pragma: no cover - defensive logging
            message = f"Execution failed: {exc}"
            self._notify("execution.error", message, level="error")
            self.progress_message = f"{label.title()} execution failed"
            return

        if execution_result is False:
            message = f"{label.title()} run cancelled."
            self._notify(
                "execution.cancelled",
                message,
                level="warning",
                reason="executor_cancelled",
            )
            self._append_status(f"{label.title()} cancelled")
            self.progress_message = f"{label.title()} cancelled"
            return

        self._notify("execution.complete", f"{label.title()} run finished.")
        self._append_status(f"{label.title()} complete")
        self.progress_message = f"{label.title()} complete"

    def _show_help(self) -> None:
        """Show help message."""
        self._append_status(
            "F1 help: arrows navigate, tab moves focus, space toggles, F10 executes plan."
        )

    def _handle_detect(self) -> None:
        """Handle detection action."""
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
        summary = summarize_inventory(inventory)
        self._notify("detect.complete", "Inventory captured.", inventory=summary)
        for line in format_inventory(inventory):
            self._append_status(f"  {line}")
        self.progress_message = "Inventory ready"

    def _handle_plan(self) -> None:
        """Handle plan building action."""
        if not self._ensure_inventory():
            return

        overrides = self._combine_overrides(None)
        overrides.setdefault("mode", "interactive")

        self.progress_message = "Planning actions..."
        self._notify(
            "plan.start",
            "Building plan from TUI.",
            overrides=dict(overrides),
        )
        self._render()
        try:
            plan_data = self.planner(self.last_inventory, overrides)
        except Exception as exc:  # pragma: no cover - defensive logging
            message = f"Plan failed: {exc}"
            self._notify("plan.error", message, level="error")
            self.progress_message = "Plan failed"
            return

        self.last_plan = plan_data
        self.last_overrides = dict(overrides)
        summary = plan_module.summarize_plan(plan_data)
        self._notify("plan.complete", "Plan ready for review.", summary=summary)
        for line in format_plan(plan_data)[:6]:
            self._append_status(f"  {line}")
        self.progress_message = "Plan ready"

    def _handle_run(self) -> None:
        """Handle plan execution action."""
        if self.last_plan is None:
            self._handle_plan()
            if self.last_plan is None:
                return

        overrides = dict(self.last_overrides or self._combine_overrides(None))
        if self.last_inventory is not None:
            overrides.setdefault("inventory", self.last_inventory)

        if not self._confirm_execution("execution", overrides):
            return

        self.progress_message = "Executing plan..."
        self._notify(
            "execution.start",
            "Executing plan from TUI.",
            overrides=dict(overrides),
        )
        self._render()
        try:
            spinner(0.2, "Preparing")
            execution_result = self.executor(self.last_plan, overrides)
        except Exception as exc:  # pragma: no cover - defensive logging
            message = f"Execution failed: {exc}"
            self._notify("execution.error", message, level="error")
            self.progress_message = "Execution failed"
            return

        if execution_result is False:
            self._notify(
                "execution.cancelled",
                "Execution cancelled before running.",
                level="warning",
                reason="executor_cancelled",
            )
            self._append_status("Execution cancelled")
            self.progress_message = "Execution cancelled"
            return

        self._notify("execution.complete", "Execution finished from TUI.")
        self._append_status("Execution complete")
        self.progress_message = "Execution complete"

    def _handle_logs(self) -> None:
        """Handle logs tab selection."""
        self._append_status("Logs tab selected")
        self.progress_message = "Viewing logs"

    def _handle_auto_all(self) -> None:
        """Handle auto scrub all action."""
        self._plan_and_optionally_execute(
            {"mode": "auto-all", "auto_all": True, "force": True},
            label="auto scrub",
            execute=True,
        )

    def _handle_cleanup_only(self) -> None:
        """Handle cleanup only action."""
        self._plan_and_optionally_execute(
            {"mode": "cleanup-only", "cleanup_only": True},
            label="cleanup only",
            execute=True,
        )

    def _handle_diagnostics(self) -> None:
        """Handle diagnostics action."""
        self._plan_and_optionally_execute(
            {"mode": "diagnose", "diagnose": True},
            label="diagnostics",
            execute=True,
        )

    def _handle_targeted(self, *, execute: bool) -> None:
        """Handle targeted scrub action."""
        if not self._ensure_inventory():
            return

        selected = self._selected_targets()
        if not selected:
            self._append_status("Select at least one Office version before running targeted scrub.")
            self.progress_message = "Target selection required"
            return

        joined = ",".join(selected)
        overrides = {"mode": f"target:{joined}", "target": joined}
        self._plan_and_optionally_execute(
            overrides,
            label="targeted scrub",
            execute=execute,
        )

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
