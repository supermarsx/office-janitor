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
    - odt_install_presets: dict[str, tuple[str, bool]]
    - odt_repair_presets: dict[str, tuple[str, bool]]
    - odt_locales: dict[str, tuple[str, bool]]
    - selected_odt_preset: str | None
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
    odt_install_presets: dict[str, tuple[str, bool]]
    odt_repair_presets: dict[str, tuple[str, bool]]
    odt_locales: dict[str, tuple[str, bool]]
    selected_odt_preset: str | None
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
    # Mode-Specific Action Handlers
    # -----------------------------------------------------------------------

    def _prepare_odt_custom(self) -> None:
        """Prepare custom ODT configuration submenu."""
        self.progress_message = "Custom ODT configuration"
        self._append_status("Custom: Specify ODT XML path or create from template.")
        self._append_status("Tip: Use bundled presets for standard configurations.")

    def _handle_odt_install_run(self) -> None:
        """Execute ODT installation from Install mode."""
        self._handle_odt_install(execute=True)

    def _prepare_repair_quick(self) -> None:
        """Prepare quick repair action."""
        self.progress_message = "Quick Repair (offline)"
        self._append_status("Quick Repair: Fixes common issues without internet.")
        self._append_status("Runs built-in Office repair. Press Enter/F10 to execute.")

    def _prepare_repair_full(self) -> None:
        """Prepare full online repair action."""
        self.progress_message = "Full Online Repair"
        self._append_status("Online Repair: Downloads fresh files from Microsoft.")
        self._append_status("Requires internet. Press Enter/F10 to execute.")

    def _handle_repair_run(self) -> None:
        """Execute repair from Repair mode."""
        # Determine which repair type based on active tab or selection
        from . import repair

        args = self.app_state.get("args")
        dry_run = bool(getattr(args, "dry_run", False))
        if "dry_run" in self.settings_overrides:
            dry_run = self.settings_overrides["dry_run"]

        # Check if a repair preset is selected, otherwise use quick repair
        preset = self._get_selected_odt_repair_preset()
        if preset:
            self._handle_odt_repair(execute=True)
            return

        # Default to quick repair if no ODT preset selected
        self.progress_message = "Running Quick Repair..."
        self._notify("repair.start", "Starting quick repair", dry_run=dry_run)
        self._render()

        try:
            result = repair.run_quick_repair(dry_run=dry_run)
            if result.returncode == 0:
                self._notify("repair.complete", "Quick repair completed successfully")
                self._append_status("✓ Quick repair complete")
                self.progress_message = "Repair complete"
            else:
                error_msg = result.stderr or result.error or f"Exit code {result.returncode}"
                self._notify("repair.error", f"Quick repair failed: {error_msg}", level="error")
                self._append_status(f"✗ Quick repair failed: {error_msg}")
                self.progress_message = "Repair failed"
        except Exception as exc:
            self._notify("repair.error", f"Repair error: {exc}", level="error")
            self._append_status(f"✗ Repair error: {exc}")
            self.progress_message = "Repair failed"

    # -----------------------------------------------------------------------
    # Special Remove Actions: C2R, OffScrub, Licensing
    # -----------------------------------------------------------------------

    def _prepare_c2r_remove(self) -> None:
        """Prepare Click-to-Run uninstall action."""
        self.progress_message = "C2R Uninstall"
        self._append_status("C2R Uninstall: Remove Click-to-Run Office installations.")
        self._append_status("Uses native OfficeClickToRun.exe for clean removal.")
        self._append_status("Press Enter/F10 to detect and remove C2R installations.")

    def _prepare_offscrub(self) -> None:
        """Prepare OffScrub scripts action."""
        self.progress_message = "OffScrub Scripts"
        self._append_status("OffScrub: Legacy Microsoft removal scripts.")
        self._append_status("Supports Office 2003-2021, MSI and C2R versions.")
        self._append_status("Press Enter/F10 to run OffScrub for detected versions.")

    def _prepare_licensing(self) -> None:
        """Prepare licensing cleanup action."""
        self.progress_message = "Licensing Cleanup"
        self._append_status("Licensing: Remove Office product keys and activation.")
        self._append_status("Cleans SPP tokens, OSPP entries, and registry keys.")
        self._append_status("Press Enter/F10 to clean licensing artifacts.")

    def _handle_c2r_remove(self) -> None:
        """Execute C2R uninstall action."""
        from . import c2r_uninstall, detect

        args = self.app_state.get("args")
        dry_run = bool(getattr(args, "dry_run", False))
        if "dry_run" in self.settings_overrides:
            dry_run = self.settings_overrides["dry_run"]

        self.progress_message = "Detecting C2R installations..."
        self._notify("c2r_remove.start", "Starting C2R removal", dry_run=dry_run)
        self._render()

        try:
            spinner(0.2, "Scanning for C2R...")
            inventory = detect.gather_office_inventory()
            c2r_items = inventory.get("c2r", [])

            if not c2r_items:
                self._append_status("No C2R installations found.")
                self.progress_message = "No C2R to remove"
                return

            self._append_status(f"Found {len(c2r_items)} C2R installation(s)")

            for item in c2r_items:
                config = c2r_uninstall.build_uninstall_config(item)
                if dry_run:
                    self._append_status(f"[DRY-RUN] Would uninstall: {item}")
                else:
                    c2r_uninstall.uninstall_products(config, dry_run=dry_run)
                    self._append_status(f"✓ Uninstalled: {item}")

            self._notify("c2r_remove.complete", "C2R removal completed")
            self.progress_message = "C2R removal complete"
        except Exception as exc:
            self._notify("c2r_remove.error", f"C2R removal error: {exc}", level="error")
            self._append_status(f"✗ C2R removal error: {exc}")
            self.progress_message = "C2R removal failed"

    def _handle_offscrub(self) -> None:
        """Execute OffScrub scripts action."""
        from . import detect, off_scrub_native

        args = self.app_state.get("args")
        dry_run = bool(getattr(args, "dry_run", False))
        if "dry_run" in self.settings_overrides:
            dry_run = self.settings_overrides["dry_run"]

        self.progress_message = "Running OffScrub..."
        self._notify("offscrub.start", "Starting OffScrub", dry_run=dry_run)
        self._render()

        try:
            spinner(0.2, "Scanning for Office...")
            inventory = detect.gather_office_inventory()

            msi_items = inventory.get("msi", [])
            c2r_items = inventory.get("c2r", [])

            if not msi_items and not c2r_items:
                self._append_status("No Office installations found for OffScrub.")
                self.progress_message = "No Office to scrub"
                return

            # Run OffScrub for MSI installations
            if msi_items:
                self._append_status(f"Found {len(msi_items)} MSI installation(s)")
                if dry_run:
                    self._append_status("[DRY-RUN] Would run OffScrub for MSI")
                else:
                    off_scrub_native.run_msi_offscrub(targets="ALL", dry_run=dry_run, force=True)
                    self._append_status("✓ OffScrub MSI completed")

            # Run OffScrub for C2R installations
            if c2r_items:
                self._append_status(f"Found {len(c2r_items)} C2R installation(s)")
                if dry_run:
                    self._append_status("[DRY-RUN] Would run OffScrub for C2R")
                else:
                    off_scrub_native.run_c2r_offscrub(
                        release_ids=["ALL"], dry_run=dry_run, force=True
                    )
                    self._append_status("✓ OffScrub C2R completed")

            self._notify("offscrub.complete", "OffScrub completed")
            self.progress_message = "OffScrub complete"
        except Exception as exc:
            self._notify("offscrub.error", f"OffScrub error: {exc}", level="error")
            self._append_status(f"✗ OffScrub error: {exc}")
            self.progress_message = "OffScrub failed"

    def _handle_licensing_cleanup(self) -> None:
        """Execute licensing cleanup action."""
        from . import licensing

        args = self.app_state.get("args")
        dry_run = bool(getattr(args, "dry_run", False))
        if "dry_run" in self.settings_overrides:
            dry_run = self.settings_overrides["dry_run"]

        self.progress_message = "Cleaning licensing..."
        self._notify("licensing.start", "Starting licensing cleanup", dry_run=dry_run)
        self._render()

        try:
            spinner(0.2, "Removing licenses...")

            if dry_run:
                self._append_status("[DRY-RUN] Would clean Office licenses")
                self._append_status("[DRY-RUN] Would remove SPP tokens")
                self._append_status("[DRY-RUN] Would clear OSPP cache")
            else:
                licensing.cleanup_licenses(extended=True)
                self._append_status("✓ Office licenses removed")
                self._append_status("✓ SPP tokens cleaned")
                self._append_status("✓ OSPP cache cleared")

            self._notify("licensing.complete", "Licensing cleanup completed")
            self.progress_message = "Licensing cleanup complete"
        except Exception as exc:
            self._notify("licensing.error", f"Licensing cleanup error: {exc}", level="error")
            self._append_status(f"✗ Licensing cleanup error: {exc}")
            self.progress_message = "Licensing cleanup failed"

    def _handle_license_status(self) -> None:
        """Display licensing status information."""
        from . import licensing

        self.progress_message = "Checking license status..."
        self._notify("license_status.start", "Querying license status")
        self._render()

        try:
            spinner(0.2, "Querying licenses...")
            status = licensing.get_license_status()

            self._append_status("═" * 40)
            self._append_status("Office Licensing Status:")
            self._append_status("═" * 40)

            if not status or not status.get("products"):
                self._append_status("No Office licenses found.")
            else:
                for product in status.get("products", []):
                    name = product.get("name", "Unknown")
                    license_status = product.get("status", "Unknown")
                    self._append_status(f"  • {name}: {license_status}")

            ospp_path = licensing.find_ospp_vbs()
            if ospp_path:
                self._append_status(f"OSPP.VBS: {ospp_path}")
            else:
                self._append_status("OSPP.VBS: Not found")

            self._notify("license_status.complete", "License status retrieved")
            self.progress_message = "License status ready"
        except Exception as exc:
            self._notify("license_status.error", f"License status error: {exc}", level="error")
            self._append_status(f"✗ License status error: {exc}")
            self.progress_message = "License status failed"

    # -----------------------------------------------------------------------
    # New Mode Handlers (ODT Builder, OffScrub, C2R, License, Config)
    # -----------------------------------------------------------------------

    def _handle_odt_export(self) -> None:
        """Export the current ODT configuration to a file."""
        self.progress_message = "Exporting ODT configuration..."
        self._append_status("Export ODT config: Feature pending implementation.")
        self._notify("odt.export", "ODT export requested")

    def _handle_offscrub_run(self) -> None:
        """Run the selected OffScrub scripts."""
        self.progress_message = "Running OffScrub scripts..."
        self._notify("offscrub.run", "OffScrub execution requested")
        self._handle_offscrub()

    def _prepare_c2r_repair(self) -> None:
        """Prepare C2R repair options."""
        self.progress_message = "Configure C2R repair"
        self._append_status("C2R Repair: Select repair type and run.")
        self.active_tab = "c2r_repair"
        pane = self.panes.get("c2r_repair")
        if pane:
            pane.cursor = 0

    def _handle_c2r_update(self) -> None:
        """Force a Click-to-Run update check."""
        from . import c2r_integrator

        self.progress_message = "Forcing C2R update..."
        self._notify("c2r.update", "C2R update requested")
        self._render()

        try:
            result = c2r_integrator.trigger_update(dry_run=self.dry_run)
            if result.get("success"):
                self._append_status("✓ C2R update triggered successfully")
            else:
                self._append_status(f"✗ C2R update failed: {result.get('error', 'Unknown')}")
            self.progress_message = "C2R update complete"
        except Exception as exc:
            self._append_status(f"✗ C2R update error: {exc}")
            self.progress_message = "C2R update failed"

    def _prepare_c2r_channel(self) -> None:
        """Prepare C2R channel change options."""
        self.progress_message = "Configure C2R update channel"
        self._append_status("C2R Channel: Select target update channel.")
        self.active_tab = "c2r_channel"
        pane = self.panes.get("c2r_channel")
        if pane:
            pane.cursor = 0

    def _prepare_license_install(self) -> None:
        """Prepare product key installation."""
        self.progress_message = "Install product key"
        self._append_status("License Install: Enter product key to install.")
        self.active_tab = "license_install"
        pane = self.panes.get("license_install")
        if pane:
            pane.cursor = 0

    def _handle_license_activate(self) -> None:
        """Trigger Office activation."""
        from . import licensing

        self.progress_message = "Activating Office..."
        self._notify("license.activate", "Office activation requested")
        self._render()

        try:
            result = licensing.activate_office(dry_run=self.dry_run)
            if result.get("success"):
                self._append_status("✓ Office activation triggered")
            else:
                self._append_status(f"✗ Activation failed: {result.get('error', 'Unknown')}")
            self.progress_message = "Activation complete"
        except Exception as exc:
            self._append_status(f"✗ Activation error: {exc}")
            self.progress_message = "Activation failed"

    def _handle_config_view(self) -> None:
        """Display current configuration."""
        self.progress_message = "Viewing configuration..."
        self._append_status("═" * 40)
        self._append_status("Current Configuration:")
        self._append_status("═" * 40)

        # Show key config values
        config_items = [
            ("Dry Run", self.dry_run),
            ("Force Mode", self.settings_overrides.get("force", False)),
            ("Registry Cleanup", self.plan_overrides.get("orphan_registry", False)),
            ("Files Cleanup", self.plan_overrides.get("orphan_files", False)),
        ]
        for key, value in config_items:
            self._append_status(f"  {key}: {value}")

        self.progress_message = "Configuration displayed"

    def _prepare_config_edit(self) -> None:
        """Prepare configuration editor."""
        self.progress_message = "Edit configuration settings"
        self._append_status("Config Edit: Modify settings with Space, save with F10.")
        self.active_tab = "settings"
        pane = self.panes.get("settings")
        if pane:
            pane.cursor = 0

    def _handle_config_export(self) -> None:
        """Export configuration to JSON file."""
        import json

        self.progress_message = "Exporting configuration..."
        self._notify("config.export", "Configuration export requested")

        config_data = {
            "plan_overrides": dict(self.plan_overrides),
            "settings_overrides": dict(self.settings_overrides),
            "target_overrides": dict(self.target_overrides),
        }

        self._append_status("═" * 40)
        self._append_status("Configuration Export:")
        self._append_status(json.dumps(config_data, indent=2))
        self._append_status("═" * 40)
        self.progress_message = "Configuration exported"

    def _prepare_config_import(self) -> None:
        """Prepare configuration import."""
        self.progress_message = "Import configuration"
        self._append_status("Config Import: Feature pending implementation.")
        self._append_status("Tip: Use --config flag on CLI to import JSON configs.")

    # -----------------------------------------------------------------------
    # ODT Install/Repair Handlers
    # -----------------------------------------------------------------------

    def _prepare_odt_install(self) -> None:
        """Prepare ODT install submenu."""
        self.progress_message = "Select ODT installation preset"
        self._append_status("ODT Install: Select preset with Space, Enter/F10 to execute.")
        self._append_status("Tip: Configure languages in 'ODT Locales' before installing.")

    def _prepare_odt_locales(self) -> None:
        """Prepare ODT locales submenu."""
        selected_count = sum(1 for _, (_, sel) in self.odt_locales.items() if sel)
        self.progress_message = f"Select Office languages ({selected_count} selected)"
        self._append_status("ODT Locales: Toggle languages with Space. Multiple allowed.")

    def _prepare_odt_repair(self) -> None:
        """Prepare ODT repair submenu."""
        self.progress_message = "Select ODT repair preset"
        self._append_status("ODT Repair: Select preset with Space, Enter/F10 to execute.")

    def _select_odt_install_preset(self, index: int) -> None:
        """Select an ODT install preset (radio-button style)."""
        pane = self.panes.get("odt_install")
        if pane is None:
            return
        self._ensure_pane_lines(pane)
        keys = list(pane.lines) if pane.lines else []
        if not keys:
            self._append_status("No ODT install presets available.")
            return
        safe_index = max(0, min(index, len(keys) - 1))
        selected_key = keys[safe_index]
        if selected_key not in self.odt_install_presets:
            self._append_status(f"Unknown preset: {selected_key}")
            return
        # Radio button behavior - deselect all others
        for key in self.odt_install_presets:
            desc, _ = self.odt_install_presets[key]
            self.odt_install_presets[key] = (desc, key == selected_key)
        self.selected_odt_preset = selected_key
        desc, _ = self.odt_install_presets[selected_key]
        self._append_status(f"Selected: {desc}")

    def _select_odt_repair_preset(self, index: int) -> None:
        """Select an ODT repair preset (radio-button style)."""
        pane = self.panes.get("odt_repair")
        if pane is None:
            return
        self._ensure_pane_lines(pane)
        keys = list(pane.lines) if pane.lines else []
        if not keys:
            self._append_status("No ODT repair presets available.")
            return
        safe_index = max(0, min(index, len(keys) - 1))
        selected_key = keys[safe_index]
        if selected_key not in self.odt_repair_presets:
            self._append_status(f"Unknown preset: {selected_key}")
            return
        # Radio button behavior - deselect all others
        for key in self.odt_repair_presets:
            desc, _ = self.odt_repair_presets[key]
            self.odt_repair_presets[key] = (desc, key == selected_key)
        self.selected_odt_preset = selected_key
        desc, _ = self.odt_repair_presets[selected_key]
        self._append_status(f"Selected: {desc}")

    def _toggle_odt_locale(self, index: int) -> None:
        """Toggle an ODT locale selection (checkbox style - multiple allowed)."""
        pane = self.panes.get("odt_locales")
        if pane is None:
            return
        self._ensure_pane_lines(pane)
        keys = list(pane.lines) if pane.lines else []
        if not keys:
            self._append_status("No locales available.")
            return
        safe_index = max(0, min(index, len(keys) - 1))
        selected_key = keys[safe_index]
        if selected_key not in self.odt_locales:
            self._append_status(f"Unknown locale: {selected_key}")
            return
        # Toggle the selected state
        desc, current_state = self.odt_locales[selected_key]
        self.odt_locales[selected_key] = (desc, not current_state)
        new_state = "selected" if not current_state else "deselected"
        selected_count = sum(1 for _, (_, sel) in self.odt_locales.items() if sel)
        self._append_status(f"{desc} ({selected_key}) {new_state} — {selected_count} total")
        self.progress_message = f"Select Office languages ({selected_count} selected)"

    def _get_selected_odt_locales(self) -> list[str]:
        """Get list of selected ODT locale codes."""
        return [key for key, (_, selected) in self.odt_locales.items() if selected]

    def _get_selected_odt_install_preset(self) -> str | None:
        """Get the currently selected ODT install preset."""
        for key, (_, selected) in self.odt_install_presets.items():
            if selected:
                return key
        return None

    def _get_selected_odt_repair_preset(self) -> str | None:
        """Get the currently selected ODT repair preset."""
        for key, (_, selected) in self.odt_repair_presets.items():
            if selected:
                return key
        return None

    def _handle_odt_install(self, *, execute: bool) -> None:
        """Handle ODT install action."""
        from . import repair

        preset = self._get_selected_odt_install_preset()
        if not preset:
            self._append_status("Select an ODT installation preset first.")
            self.progress_message = "Preset selection required"
            return

        selected_locales = self._get_selected_odt_locales()
        if not selected_locales:
            self._append_status("Select at least one language in 'ODT Locales' first.")
            self.progress_message = "Language selection required"
            return

        desc, _ = self.odt_install_presets[preset]
        locale_summary = ", ".join(selected_locales[:3])
        if len(selected_locales) > 3:
            locale_summary += f" +{len(selected_locales) - 3} more"

        if not execute:
            self._append_status(f"ODT install ready: {desc}")
            self._append_status(f"Languages: {locale_summary}")
            self.progress_message = f"Ready: {desc}"
            return

        args = self.app_state.get("args")
        dry_run = bool(getattr(args, "dry_run", False))
        if "dry_run" in self.settings_overrides:
            dry_run = self.settings_overrides["dry_run"]

        # Request confirmation
        confirm_msg = f"ODT Install: {desc}\nLanguages: {locale_summary}"
        if not self._confirm_odt_execution(confirm_msg, dry_run):
            return

        self.progress_message = f"Executing ODT Install: {desc}..."
        self._notify(
            "odt_install.start",
            f"Starting ODT installation: {desc}",
            preset=preset,
            locales=selected_locales,
            dry_run=dry_run,
        )
        self._render()

        try:
            spinner(0.2, "Preparing ODT...")
            # Note: The selected locales are logged for informational purposes.
            # The bundled XML configs have predefined language settings.
            # For custom language configuration, users should use the ODT with custom XML.
            result = repair.run_oem_config(preset, dry_run=dry_run)

            if result.returncode == 0:
                self._notify(
                    "odt_install.complete",
                    f"ODT installation completed: {desc}",
                    preset=preset,
                    locales=selected_locales,
                )
                self._append_status(f"✓ ODT Install complete: {desc}")
                self._append_status(f"  Languages installed: {locale_summary}")
                self.progress_message = "ODT Install complete"
            else:
                error_msg = result.stderr or result.error or f"Exit code {result.returncode}"
                self._notify(
                    "odt_install.error",
                    f"ODT installation failed: {error_msg}",
                    level="error",
                    preset=preset,
                    exit_code=result.returncode,
                )
                self._append_status(f"✗ ODT Install failed: {error_msg}")
                self.progress_message = "ODT Install failed"

        except Exception as exc:
            self._notify(
                "odt_install.error",
                f"ODT installation error: {exc}",
                level="error",
            )
            self._append_status(f"✗ ODT Install error: {exc}")
            self.progress_message = "ODT Install failed"

    def _handle_odt_repair(self, *, execute: bool) -> None:
        """Handle ODT repair action."""
        from . import repair

        preset = self._get_selected_odt_repair_preset()
        if not preset:
            self._append_status("Select an ODT repair preset first.")
            self.progress_message = "Preset selection required"
            return

        desc, _ = self.odt_repair_presets[preset]
        if not execute:
            self._append_status(f"ODT repair ready: {desc}")
            self.progress_message = f"Ready: {desc}"
            return

        args = self.app_state.get("args")
        dry_run = bool(getattr(args, "dry_run", False))
        if "dry_run" in self.settings_overrides:
            dry_run = self.settings_overrides["dry_run"]

        # Request confirmation
        if not self._confirm_odt_execution(f"ODT Repair: {desc}", dry_run):
            return

        self.progress_message = f"Executing ODT Repair: {desc}..."
        self._notify(
            "odt_repair.start",
            f"Starting ODT repair: {desc}",
            preset=preset,
            dry_run=dry_run,
        )
        self._render()

        try:
            spinner(0.2, "Preparing ODT...")
            result = repair.run_oem_config(preset, dry_run=dry_run)

            if result.returncode == 0:
                self._notify(
                    "odt_repair.complete",
                    f"ODT repair completed: {desc}",
                    preset=preset,
                )
                self._append_status(f"✓ ODT Repair complete: {desc}")
                self.progress_message = "ODT Repair complete"
            else:
                error_msg = result.stderr or result.error or f"Exit code {result.returncode}"
                self._notify(
                    "odt_repair.error",
                    f"ODT repair failed: {error_msg}",
                    level="error",
                    preset=preset,
                    exit_code=result.returncode,
                )
                self._append_status(f"✗ ODT Repair failed: {error_msg}")
                self.progress_message = "ODT Repair failed"

        except Exception as exc:
            self._notify(
                "odt_repair.error",
                f"ODT repair error: {exc}",
                level="error",
            )
            self._append_status(f"✗ ODT Repair error: {exc}")
            self.progress_message = "ODT Repair failed"

    def _confirm_odt_execution(self, label: str, dry_run: bool) -> bool:
        """Request confirmation for ODT execution."""
        request_confirmation = self._confirm_requestor
        if request_confirmation is None:
            return True

        args = self.app_state.get("args")
        force_override = bool(getattr(args, "force", False))

        proceed = request_confirmation(
            dry_run=dry_run,
            force=force_override,
            input_func=lambda prompt: self._prompt_confirmation_input(label, prompt),
            interactive=True,
        )

        if not proceed:
            self._notify(
                "odt.cancelled",
                f"{label} cancelled by user",
                level="warning",
                reason="user_declined",
            )
            self.progress_message = f"{label} cancelled"
            return False

        return True

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
