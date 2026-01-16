"""!
@file tui_actions.py
@brief Action handler methods for the Office Janitor TUI.

@details Contains the TUIActionsMixin class that provides action handlers
for detection, planning, execution, and navigation. This mixin is designed
to be combined with the main OfficeJanitorTUI class.
"""

from __future__ import annotations

from collections.abc import Mapping, MutableMapping
from typing import TYPE_CHECKING, Callable

from . import plan as plan_module
from .tui_helpers import (
    format_inventory,
    format_plan,
    read_input_line,
    spinner,
    summarize_inventory,
)

if TYPE_CHECKING:
    from collections import deque

    from .tui import PaneContext


class TUIActionsMixin:
    """!
    @brief Mixin providing action handlers for TUI operations.
    @details This mixin expects the following attributes on the class:
    - app_state: MutableMapping
    - detector: Callable
    - planner: Callable
    - executor: Callable
    - _confirm_requestor: Callable | None
    - last_inventory: Mapping | None
    - last_plan: list | None
    - last_overrides: dict | None
    - progress_message: str
    - panes: dict[str, PaneContext]
    - plan_overrides: dict[str, bool]
    - target_overrides: dict[str, bool]
    - settings_overrides: dict[str, bool]
    - status_lines: list[str]
    - compact_layout: bool
    - _running: bool
    """

    # Type hints for mixin attributes (defined in OfficeJanitorTUI)
    app_state: MutableMapping[str, object]
    detector: Callable[[], Mapping[str, object]]
    planner: Callable[..., object]
    executor: Callable[..., object]
    _confirm_requestor: Callable[..., bool] | None
    last_inventory: Mapping[str, object] | None
    last_plan: list[dict[str, object]] | None
    last_overrides: dict[str, object] | None
    progress_message: str
    panes: dict[str, object]
    plan_overrides: dict[str, bool]
    target_overrides: dict[str, bool]
    settings_overrides: dict[str, bool]
    status_lines: list[str]
    compact_layout: bool
    _running: bool
    human_logger: object
    machine_logger: object
    event_queue: deque[dict[str, object]]
    emit_event: object

    # Methods that must be provided by the main class
    def _render(self) -> None:
        raise NotImplementedError  # pragma: no cover

    def _append_status(self, message: str) -> None:
        raise NotImplementedError  # pragma: no cover

    def _notify(self, event: str, message: str, *, level: str = "info", **payload: object) -> None:
        raise NotImplementedError  # pragma: no cover

    def _ensure_pane_lines(self, pane: PaneContext) -> list[tuple[str, str]]:
        raise NotImplementedError  # pragma: no cover

    def _get_pane_filter(self, pane_name: str) -> str:
        raise NotImplementedError  # pragma: no cover

    def _set_pane_filter(self, pane_name: str, value: str) -> None:
        raise NotImplementedError  # pragma: no cover

    def _filter_entries(
        self, pane: PaneContext, entries: list[tuple[str, str]]
    ) -> list[tuple[str, str]]:
        raise NotImplementedError  # pragma: no cover

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
    # Confirmation
    # -----------------------------------------------------------------------

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

    def _prompt_confirmation_input(self, label: str, prompt: str) -> str:
        """!
        @brief Display a confirmation prompt and collect a response inside the TUI.
        """
        raise NotImplementedError  # pragma: no cover - implemented in main class

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
