"""!
@brief Text-based user interface (TUI) engine.
@details Implements the ANSI/VT driven interface described in the project
specification.  The implementation keeps dependencies to the standard library
only while providing a co-operative event loop that drains orchestrator
progress events and handles keyboard commands.  The layout follows a header,
navigation column, and tabbed content panes so that additional widgets can be
rendered predictably across platforms.
"""

from __future__ import annotations

import os
import platform
import sys
import time
from collections import deque
from collections.abc import Mapping, MutableMapping
from dataclasses import dataclass, field
from typing import Any, Callable

try:  # pragma: no cover - Windows specific
    import msvcrt as _msvcrt  # type: ignore[import-not-found]
except ImportError:  # pragma: no cover - non-Windows hosts
    _msvcrt = None  # type: ignore[assignment]

msvcrt: Any = _msvcrt

from . import constants, version
from . import plan as plan_module


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
        self.planner: Callable[
            [Mapping[str, object], Mapping[str, object] | None], list[dict]
        ] = self.app_state[
            "planner"
        ]  # type: ignore[assignment]
        self.executor: Callable[
            [list[dict], Mapping[str, object] | None], bool | None
        ] = self.app_state[
            "executor"
        ]  # type: ignore[assignment]
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
        self.last_plan: list[dict] | None = None
        self.status_lines: list[str] = []
        self.progress_message = "Ready"
        self.log_lines: list[str] = []
        self.ansi_supported = _supports_ansi() and not bool(
            getattr(self.app_state.get("args"), "no_color", False)
        )
        self._key_reader: Callable[[], str] | None = self.app_state.get(
            "key_reader"
        )  # type: ignore[assignment]
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

        if not self.ansi_supported and not _supports_ansi():
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

    def _ensure_pane_lines(self, pane: PaneContext) -> list[tuple[str, str]]:
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
            entries = [(line, line) for line in _format_inventory(self.last_inventory)]
        elif pane.name in {"auto", "cleanup", "diagnostics"}:
            entries = []
        else:
            entries = [(line, line) for line in pane.lines]

        return self._filter_entries(pane, entries)

    def _get_pane_filter(self, pane_name: str) -> str:
        return self.list_filters.get(pane_name, "")

    def _set_pane_filter(self, pane_name: str, value: str) -> None:
        normalized = value.strip()
        if normalized:
            self.list_filters[pane_name] = normalized
        else:
            self.list_filters.pop(pane_name, None)

    def _filter_entries(
        self, pane: PaneContext, entries: list[tuple[str, str]]
    ) -> list[tuple[str, str]]:
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
        label = pane.name.replace("_", " ").title()
        previous_filter = self._get_pane_filter(pane.name)
        previous_message = self.progress_message
        response = ""
        try:
            self.progress_message = f"Filter {label}"
            self._render()
            try:
                response = _read_input_line(f"Filter {label} (empty to clear): ")
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

    def _move_nav(self, offset: int) -> None:
        size = len(self.navigation)
        self.nav_index = (self.nav_index + offset) % size
        self.active_tab = self.navigation[self.nav_index].name
        self.focus_area = "nav"

    def _activate_nav(self) -> None:
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
        reader = self._key_reader or _default_key_reader
        try:
            raw = reader()
        except StopIteration:
            self._running = False
            return "quit"
        except Exception:
            return ""
        return _decode_key(raw)

    def _render(self) -> None:
        width = 80 if self.compact_layout else 96
        left_width = 24 if self.compact_layout else 28
        _clear_screen()
        sys.stdout.write(self._render_header(width) + "\n")
        sys.stdout.write(_divider(width) + "\n")

        nav_lines = self._render_navigation(left_width)
        content_lines = self._render_content(width - left_width - 1)

        max_lines = max(len(nav_lines), len(content_lines))
        for index in range(max_lines):
            left_text = nav_lines[index] if index < len(nav_lines) else ""
            right_text = content_lines[index] if index < len(content_lines) else ""
            sys.stdout.write(
                f"{left_text.ljust(left_width)} {right_text[: width - left_width - 1]}\n"
            )

        sys.stdout.write(_divider(width) + "\n")
        sys.stdout.write(self._render_footer() + "\n")
        sys.stdout.flush()

    def _render_header(self, width: int) -> str:
        metadata = version.build_info()
        node = platform.node() or os.environ.get("COMPUTERNAME", "Unknown")
        header = f"Office Janitor {metadata['version']}"
        header += f" — {self.progress_message}"
        header += f" — {node}"
        return header[:width]

    def _render_navigation(self, width: int) -> list[str]:
        lines: list[str] = ["Navigation:"]
        for index, item in enumerate(self.navigation):
            prefix = "➤" if index == self.nav_index else " "
            if not self.ansi_supported or index != self.nav_index:
                line = f"{prefix} {item.label}"
            else:
                line = f"\x1b[7m{prefix} {item.label}\x1b[0m"
            lines.append(line[:width])
        lines.append("")
        lines.append("Status log:")
        lines.extend(self.status_lines[-(12 if self.compact_layout else 18) :])
        return lines

    def _render_content(self, width: int) -> list[str]:
        if self.active_tab == "detect":
            return self._render_inventory_pane(width)
        if self.active_tab == "plan":
            return self._render_plan_pane(width)
        if self.active_tab == "targeted":
            return self._render_targeted_pane(width)
        if self.active_tab == "auto":
            return self._render_auto_pane(width)
        if self.active_tab == "cleanup":
            return self._render_cleanup_pane(width)
        if self.active_tab == "diagnostics":
            return self._render_diagnostics_pane(width)
        if self.active_tab == "run":
            return self._render_run_pane(width)
        if self.active_tab == "logs":
            return self._render_logs_pane(width)
        if self.active_tab == "settings":
            return self._render_settings_pane(width)
        return ["Select an option with Enter"]

    def _render_footer(self) -> str:
        help_text = (
            "Arrows navigate • Tab switches focus • Space toggles • Enter confirms • "
            "F10 run • / filter • F1 help • Q quits"
        )
        return help_text

    def _render_inventory_pane(self, width: int) -> list[str]:
        lines = ["Inventory summary:"]
        if self.last_inventory is None:
            lines.append("No inventory collected yet.")
        else:
            pane = self.panes["detect"]
            entries = self._ensure_pane_lines(pane)
            active_filter = self._get_pane_filter(pane.name)
            if active_filter:
                lines.append(f"Filter: {active_filter}")
            if entries:
                lines.extend(text for _, text in entries)
            else:
                lines.append("No matching entries.")
        pane = self.panes["detect"]
        if self.last_inventory is None:
            pane.lines = []
        return [line[:width] for line in lines]

    def _render_plan_pane(self, width: int) -> list[str]:
        lines = ["Plan options:"]
        pane = self.panes["plan"]
        entries = self._ensure_pane_lines(pane)
        active_filter = self._get_pane_filter(pane.name)
        if active_filter:
            lines.append(f"Filter: {active_filter}")
        if entries:
            for index, (_, label) in enumerate(entries):
                cursor = "➤" if pane.cursor == index else " "
                lines.append(f"{cursor} {label}")
        else:
            lines.append("No matching options.")
        lines.append("")
        summary = _format_plan(self.last_plan)
        lines.extend(summary)
        return [line[:width] for line in lines]

    def _render_targeted_pane(self, width: int) -> list[str]:
        lines = ["Targeted scrub targets:"]
        pane = self.panes["targeted"]
        entries = self._ensure_pane_lines(pane)
        active_filter = self._get_pane_filter(pane.name)
        if active_filter:
            lines.append(f"Filter: {active_filter}")
        if entries:
            for index, (_, label) in enumerate(entries):
                cursor = "➤" if pane.cursor == index else " "
                lines.append(f"{cursor} {label}")
        else:
            lines.append("No matching targets.")
        lines.append("")
        lines.append("Select versions with Space, Enter/F10 to run targeted scrub.")
        if self.last_inventory is None:
            lines.append("Inventory not captured yet; selecting targets will prompt detection.")
        return [line[:width] for line in lines]

    def _render_auto_pane(self, width: int) -> list[str]:
        lines = [
            "Auto scrub:",
            "Run a full detection and removal pass for all detected Office versions.",
            "Press Enter or F10 to start the auto scrub run.",
        ]
        return [line[:width] for line in lines]

    def _render_cleanup_pane(self, width: int) -> list[str]:
        lines = [
            "Cleanup only:",
            "Removes residue such as licenses and scheduled tasks without uninstalling suites.",
            "Press Enter or F10 to execute cleanup steps.",
        ]
        return [line[:width] for line in lines]

    def _render_diagnostics_pane(self, width: int) -> list[str]:
        lines = [
            "Diagnostics only:",
            "Exports inventory and action plans without running uninstall steps.",
            "Press Enter or F10 to generate diagnostics artifacts.",
        ]
        return [line[:width] for line in lines]

    def _render_run_pane(self, width: int) -> list[str]:
        lines = ["Execution progress:"]
        pane = self.panes["run"]
        entries = self._ensure_pane_lines(pane)
        active_filter = self._get_pane_filter(pane.name)
        if active_filter:
            lines.append(f"Filter: {active_filter}")
        if entries:
            lines.extend(text for _, text in entries)
        else:
            lines.append("No matching progress messages." if active_filter else "No progress yet.")
        return [line[:width] for line in lines]

    def _render_logs_pane(self, width: int) -> list[str]:
        lines = ["Log tail:"]
        if not self.log_lines:
            lines.append("No log entries yet.")
        else:
            pane = self.panes["logs"]
            entries = self._ensure_pane_lines(pane)
            active_filter = self._get_pane_filter(pane.name)
            if active_filter:
                lines.append(f"Filter: {active_filter}")
            if entries:
                start = max(0, len(entries) - 10)
                lines.extend(text for _, text in entries[start:])
            else:
                lines.append("No matching log entries.")
        return [line[:width] for line in lines]

    def _render_settings_pane(self, width: int) -> list[str]:
        lines = ["Settings toggles:"]
        pane = self.panes["settings"]
        entries = self._ensure_pane_lines(pane)
        active_filter = self._get_pane_filter(pane.name)
        if active_filter:
            lines.append(f"Filter: {active_filter}")
        if entries:
            for index, (_, label) in enumerate(entries):
                cursor = "➤" if pane.cursor == index else " "
                lines.append(f"{cursor} {label}")
        else:
            lines.append("No matching settings.")
        args = self.app_state.get("args")
        lines.append("")
        lines.append(f"Log directory: {getattr(args, 'logdir', '(default)')}")
        lines.append(f"Backup directory: {getattr(args, 'backup', '(disabled)')}")
        timeout = getattr(args, "timeout", None)
        lines.append(f"Timeout: {timeout if timeout is not None else '(default)'}")
        return [line[:width] for line in lines]

    def _toggle_plan_option(self, index: int) -> None:
        pane = self.panes.get("plan")
        keys = (
            list(pane.lines)
            if pane is not None and pane.lines
            else list(self.plan_overrides.keys())
        )
        if not keys:
            return
        safe_index = max(0, min(index, len(keys) - 1))
        key = keys[safe_index]
        self.plan_overrides[key] = not self.plan_overrides[key]
        self._append_status(f"Plan option '{key}' set to {self.plan_overrides[key]}")

    def _toggle_target(self, index: int) -> None:
        pane = self.panes.get("targeted")
        keys = (
            list(pane.lines)
            if pane is not None and pane.lines
            else list(self.target_overrides.keys())
        )
        if not keys:
            return
        safe_index = max(0, min(index, len(keys) - 1))
        key = keys[safe_index]
        self.target_overrides[key] = not self.target_overrides[key]
        state = "selected" if self.target_overrides[key] else "cleared"
        self._append_status(f"Target version {key} {state}.")

    def _toggle_setting(self, index: int) -> None:
        pane = self.panes.get("settings")
        keys = (
            list(pane.lines)
            if pane is not None and pane.lines
            else list(self.settings_overrides.keys())
        )
        if not keys:
            return
        safe_index = max(0, min(index, len(keys) - 1))
        key = keys[safe_index]
        self.settings_overrides[key] = not self.settings_overrides[key]
        self._append_status(f"Setting '{key}' toggled to {self.settings_overrides[key]}")

    def _collect_settings_overrides(self) -> dict[str, object]:
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
        combined: dict[str, object] = {}
        combined.update(self._collect_settings_overrides())
        combined.update(self._collect_plan_overrides())
        if extra:
            combined.update({key: extra[key] for key in extra})
        return combined

    def _selected_targets(self) -> list[str]:
        return [version for version, enabled in self.target_overrides.items() if enabled]

    def _ensure_inventory(self) -> bool:
        if self.last_inventory is not None:
            return True
        self._handle_detect()
        return self.last_inventory is not None

    def _prepare_targeted(self) -> None:
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
        for line in _format_plan(plan_data)[:6]:
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
            _spinner(0.2, "Preparing")
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
        self._append_status(
            "F1 help: arrows navigate, tab moves focus, space toggles, F10 executes plan."
        )

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
        for line in _format_plan(plan_data)[:6]:
            self._append_status(f"  {line}")
        self.progress_message = "Plan ready"

    def _handle_run(self) -> None:
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
            _spinner(0.2, "Preparing")
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
        self._append_status("Logs tab selected")
        self.progress_message = "Viewing logs"

    def _handle_auto_all(self) -> None:
        self._plan_and_optionally_execute(
            {"mode": "auto-all", "auto_all": True},
            label="auto scrub",
            execute=True,
        )

    def _handle_cleanup_only(self) -> None:
        self._plan_and_optionally_execute(
            {"mode": "cleanup-only", "cleanup_only": True},
            label="cleanup only",
            execute=True,
        )

    def _handle_diagnostics(self) -> None:
        self._plan_and_optionally_execute(
            {"mode": "diagnose", "diagnose": True},
            label="diagnostics",
            execute=True,
        )

    def _handle_targeted(self, *, execute: bool) -> None:
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

    def _append_status(self, message: str) -> None:
        if self.status_lines and self.status_lines[-1] == message:
            return
        self.status_lines.append(message)
        limit = 24 if self.compact_layout else 32
        if len(self.status_lines) > limit:
            self.status_lines[:] = self.status_lines[-limit:]

    def _notify(self, event: str, message: str, *, level: str = "info", **payload: object) -> None:
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
                    self.last_plan = plan_data  # type: ignore[assignment]
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
        self.log_lines.append(line)
        limit = 200
        if len(self.log_lines) > limit:
            self.log_lines[:] = self.log_lines[-limit:]


def run_tui(app_state: Mapping[str, object]) -> None:
    """!
    @brief Convenience wrapper to create and run the TUI controller.
    """

    OfficeJanitorTUI(app_state).run()


def _supports_ansi(stream: object | None = None) -> bool:
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


def _decode_key(raw: str) -> str:
    """!
    @brief Convert raw key input into a normalized command token.
    """

    if not raw:
        return ""
    mapping = {
        "\t": "tab",
        "\r": "enter",
        "\n": "enter",
        " ": "space",
        "q": "quit",
        "Q": "quit",
        "\x1b": "escape",
    }
    if raw in mapping:
        return mapping[raw]
    if raw.startswith("\x1b"):
        sequences = {
            "\x1b[A": "up",
            "\x1b[B": "down",
            "\x1b[C": "right",
            "\x1b[D": "left",
            "\x1b[5~": "page_up",
            "\x1b[6~": "page_down",
            "\x1bOP": "f1",
            "\x1b[21~": "f10",
            "\x1b[1;9A": "up",
            "\x1b[1;9B": "down",
        }
        return sequences.get(raw, "")
    windows_sequences = {
        "\x00H": "up",
        "\x00P": "down",
        "\x00K": "left",
        "\x00M": "right",
        "\xe0H": "up",
        "\xe0P": "down",
        "\xe0K": "left",
        "\xe0M": "right",
        "\x00I": "page_up",
        "\x00Q": "page_down",
        "\xe0I": "page_up",
        "\xe0Q": "page_down",
        "\x00;": "f1",
        "\x00h": "f10",
    }
    if raw in windows_sequences:
        return windows_sequences[raw]
    trimmed = raw.strip()
    if trimmed == "/":
        return "/"
    return trimmed


def _clear_screen() -> None:
    sys.stdout.write("\x1b[2J\x1b[H")
    sys.stdout.flush()


def _divider(width: int) -> str:
    return "-" * width


def _format_inventory(inventory: Mapping[str, object]) -> list[str]:
    lines: list[str] = []
    for key, value in inventory.items():
        lines.extend(_flatten_inventory_entry(str(key), value))
    if not lines:
        lines.append("No data")
    return lines


def _flatten_inventory_entry(prefix: str, value: object) -> list[str]:
    if isinstance(value, Mapping):
        if not value:
            return [f"{prefix}: (empty)"]
        lines: list[str] = []
        for subkey, subvalue in value.items():
            nested_prefix = f"{prefix}.{subkey}"
            lines.extend(_flatten_inventory_entry(nested_prefix, subvalue))
        return lines
    if isinstance(value, (list, tuple, set)):
        items = list(value)
        if not items:
            return [f"{prefix}: (empty)"]
        return [f"{prefix}: {_stringify_inventory_value(item)}" for item in items]
    return [f"{prefix}: {_stringify_inventory_value(value)}"]


def _stringify_inventory_value(value: object) -> str:
    if isinstance(value, Mapping):
        if not value:
            return "{}"
        parts = [f"{key}={_stringify_inventory_value(subvalue)}" for key, subvalue in value.items()]
        return ", ".join(parts)
    if isinstance(value, (list, tuple, set)):
        if not value:
            return "[]"
        return ", ".join(_stringify_inventory_value(item) for item in value)
    return str(value)


def _summarize_inventory(inventory: Mapping[str, object]) -> dict[str, int]:
    summary: dict[str, int] = {}
    for key, items in inventory.items():
        try:
            count = len(items)  # type: ignore[arg-type]
        except TypeError:
            count = len(list(items))  # type: ignore[arg-type]
        summary[str(key)] = count
    return summary


def _format_plan(plan_data: list[dict] | None) -> list[str]:
    if not plan_data:
        return ["Plan not created"]

    summary = plan_module.summarize_plan(plan_data)
    lines: list[str] = [
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
            first = msvcrt.getwch()
            if first in {"\x00", "\xe0"}:
                second = msvcrt.getwch()
                return first + second
            if first == "\x1b" and msvcrt.kbhit():
                second = msvcrt.getwch()
                if second == "[":
                    third = msvcrt.getwch()
                    return first + second + third
            return first
        return ""
    except Exception:
        return _read_input_line("Command (arrows navigate, enter selects, q to quit): ").strip()
