"""!
@file tui_render.py
@brief Rendering mixins and methods for the Office Janitor TUI.

@details Contains the TUIRendererMixin class that provides all rendering
methods for panes, navigation, header/footer, and content areas. This mixin
is designed to be combined with the main OfficeJanitorTUI class.
"""

from __future__ import annotations

import os
import platform
import sys
from collections.abc import Mapping
from typing import TYPE_CHECKING

from . import version
from .tui_helpers import (
    clear_screen,
    divider,
    format_plan,
    strip_ansi,
)

if TYPE_CHECKING:
    from .tui import PaneContext


class TUIRendererMixin:
    """!
    @brief Mixin providing rendering methods for TUI panes and layout.
    @details This mixin expects the following attributes on the class:
    - ansi_supported: bool
    - compact_layout: bool
    - progress_message: str
    - navigation: list[NavigationItem]
    - nav_index: int
    - active_tab: str
    - panes: dict[str, PaneContext]
    - status_lines: list[str]
    - log_lines: list[str]
    - last_inventory: Mapping | None
    - last_plan: list | None
    - plan_overrides: dict[str, bool]
    - target_overrides: dict[str, bool]
    - settings_overrides: dict[str, bool]
    - odt_install_presets: dict[str, tuple[str, bool]]
    - odt_repair_presets: dict[str, tuple[str, bool]]
    - odt_locales: dict[str, tuple[str, bool]]
    - app_state: MutableMapping
    - current_mode: str | None
    - mode_options: list[tuple[str, str, str]]
    - mode_index: int
    """

    # Type hints for mixin attributes (defined in OfficeJanitorTUI)
    ansi_supported: bool
    compact_layout: bool
    progress_message: str
    navigation: list[str]
    nav_index: int
    active_tab: str
    panes: dict[str, object]
    status_lines: list[str]
    log_lines: list[str]
    last_inventory: Mapping[str, object] | None
    last_plan: list[dict[str, object]] | None
    plan_overrides: dict[str, bool]
    target_overrides: dict[str, bool]
    settings_overrides: dict[str, bool]
    odt_install_presets: dict[str, tuple[str, bool]]
    odt_repair_presets: dict[str, tuple[str, bool]]
    odt_locales: dict[str, tuple[str, bool]]
    app_state: dict[str, object]
    list_filters: dict[str, str]
    current_mode: str | None
    mode_options: list[tuple[str, str, str]]
    mode_index: int

    def _render(self) -> None:
        """!
        @brief Render the full TUI screen.
        """
        width = 80 if self.compact_layout else 96
        left_width = 24 if self.compact_layout else 28
        clear_screen()
        sys.stdout.write(self._render_header(width) + "\n")
        sys.stdout.write(divider(width) + "\n")

        # Mode selection screen uses full-width centered layout
        if self.current_mode is None:
            mode_lines = self._render_mode_selection(width)
            for line in mode_lines:
                sys.stdout.write(line + "\n")
        else:
            nav_lines = self._render_navigation(left_width)
            content_lines = self._render_content(width - left_width - 1)

            max_lines = max(len(nav_lines), len(content_lines))
            for index in range(max_lines):
                left_text = nav_lines[index] if index < len(nav_lines) else ""
                right_text = content_lines[index] if index < len(content_lines) else ""
                # Calculate visible length (excluding ANSI codes) for proper padding
                visible_len = len(strip_ansi(left_text))
                padding = " " * max(0, left_width - visible_len)
                sys.stdout.write(f"{left_text}{padding} {right_text[: width - left_width - 1]}\n")

        sys.stdout.write(divider(width) + "\n")
        sys.stdout.write(self._render_footer() + "\n")
        sys.stdout.flush()

    def _render_mode_selection(self, width: int) -> list[str]:
        """Render the mode selection screen."""
        lines: list[str] = []
        lines.append("")
        title = "Select Operation Mode"
        lines.append(title.center(width))
        lines.append(("═" * len(title)).center(width))
        lines.append("")

        for index, (mode_id, label, description) in enumerate(self.mode_options):
            prefix = "►" if index == self.mode_index else " "
            visible_text = f"{prefix} [{mode_id[0].upper()}] {label}"

            if self.ansi_supported and index == self.mode_index:
                # Highlight selected option
                padded = visible_text.ljust(width)
                line = f"\x1b[7m{padded}\x1b[0m"
            else:
                line = visible_text
            lines.append(line)
            # Add description under the label (dimmed if ANSI supported)
            desc_text = f"      {description}"
            if self.ansi_supported:
                desc_line = f"\x1b[2m{desc_text[:width]}\x1b[0m"
            else:
                desc_line = desc_text[:width]
            lines.append(desc_line)

        lines.append("")
        nav_help = "↑↓ Navigate  →/Enter Select  ←/Esc Back  Q Quit  F1 Help"
        lines.append(nav_help.center(width))
        lines.append("")

        # Add status log at bottom
        lines.append("Status:")
        lines.extend(self.status_lines[-(8 if self.compact_layout else 10) :])

        return lines

    def _render_header(self, width: int) -> str:
        """Render the header line with version and status."""
        metadata = version.build_info()
        node = platform.node() or os.environ.get("COMPUTERNAME", "Unknown")
        header = f"Office Janitor {metadata['version']}"
        header += f" — {self.progress_message}"
        header += f" — {node}"
        return header[:width]

    def _render_navigation(self, width: int) -> list[str]:
        """Render the navigation column."""
        mode_label = self.current_mode.title() if self.current_mode else "Mode"
        lines: list[str] = [f"{mode_label} Menu:"]
        for index, item in enumerate(self.navigation):
            prefix = "►" if index == self.nav_index else " "
            # Build visible text first, truncate to width, then apply ANSI
            visible_text = f"{prefix} {item.label}"
            truncated = visible_text[:width]
            if self.ansi_supported and index == self.nav_index:
                # Pad to width so highlight fills the column, then apply reverse video
                padded = truncated.ljust(width)
                line = f"\x1b[7m{padded}\x1b[0m"
            else:
                line = truncated
            lines.append(line)
        lines.append("")
        lines.append("Status:")
        lines.extend(self.status_lines[-(12 if self.compact_layout else 18) :])
        return lines

    def _render_content(self, width: int) -> list[str]:
        """Dispatch to the appropriate pane renderer based on active tab."""
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
        if self.active_tab == "odt_install":
            return self._render_odt_install_pane(width)
        if self.active_tab == "odt_locales":
            return self._render_odt_locales_pane(width)
        if self.active_tab == "odt_repair":
            return self._render_odt_repair_pane(width)
        if self.active_tab == "c2r_remove":
            return self._render_c2r_remove_pane(width)
        if self.active_tab == "offscrub":
            return self._render_offscrub_pane(width)
        if self.active_tab == "licensing":
            return self._render_licensing_pane(width)
        if self.active_tab == "license_status":
            return self._render_license_status_pane(width)
        if self.active_tab == "run":
            return self._render_run_pane(width)
        if self.active_tab == "logs":
            return self._render_logs_pane(width)
        if self.active_tab == "settings":
            return self._render_settings_pane(width)
        return ["Select an option with Enter"]

    def _render_footer(self) -> str:
        """Render the footer help text."""
        if self.current_mode is None:
            help_text = "↑↓ Navigate  →/Enter Select  Q Quit  F1 Help"
        else:
            help_text = (
                "↑↓ Navigate  →/Enter Select  ← Back  Tab Focus  "
                "Space Toggle  F10 Run  / Filter  Q Quit"
            )
        return help_text

    def _render_inventory_pane(self, width: int) -> list[str]:
        """Render the inventory/detect pane."""
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
        """Render the plan options pane."""
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
        summary = format_plan(self.last_plan)
        lines.extend(summary)
        return [line[:width] for line in lines]

    def _render_targeted_pane(self, width: int) -> list[str]:
        """Render the targeted scrub pane."""
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
        """Render the auto scrub pane."""
        lines = [
            "Auto scrub:",
            "Run a full detection and removal pass for all detected Office versions.",
            "Press Enter or F10 to start the auto scrub run.",
        ]
        return [line[:width] for line in lines]

    def _render_cleanup_pane(self, width: int) -> list[str]:
        """Render the cleanup only pane."""
        lines = [
            "Cleanup only:",
            "Removes residue such as licenses and scheduled tasks without uninstalling suites.",
            "Press Enter or F10 to execute cleanup steps.",
        ]
        return [line[:width] for line in lines]

    def _render_diagnostics_pane(self, width: int) -> list[str]:
        """Render the diagnostics pane."""
        lines = [
            "Diagnostics only:",
            "Exports inventory and action plans without running uninstall steps.",
            "Press Enter or F10 to generate diagnostics artifacts.",
        ]
        return [line[:width] for line in lines]

    def _render_odt_install_pane(self, width: int) -> list[str]:
        """Render the ODT installation presets pane."""
        lines = ["ODT Installation Presets:"]
        lines.append("")
        pane = self.panes["odt_install"]
        entries = self._ensure_pane_lines(pane)
        active_filter = self._get_pane_filter(pane.name)
        if active_filter:
            lines.append(f"Filter: {active_filter}")
        if entries:
            for index, (_, label) in enumerate(entries):
                cursor = "➤" if pane.cursor == index else " "
                lines.append(f"{cursor} {label}")
        else:
            lines.append("No matching presets.")
        lines.append("")
        lines.append("─" * min(width - 2, 50))
        lines.append("Select a preset with Space, Enter/F10 to install.")
        lines.append("")
        # Show selected locales summary
        selected_locales = [k for k, (_, sel) in self.odt_locales.items() if sel]
        if selected_locales:
            locale_str = ", ".join(selected_locales[:5])
            if len(selected_locales) > 5:
                locale_str += f" +{len(selected_locales) - 5} more"
            lines.append(f"Languages: {locale_str}")
        else:
            lines.append("Languages: None selected (configure in ODT Locales)")
        lines.append("")
        lines.append("Note: Configure languages in 'ODT Locales' menu.")
        return [line[:width] for line in lines]

    def _render_odt_locales_pane(self, width: int) -> list[str]:
        """Render the ODT locale selection pane."""
        lines = ["ODT Language Selection:"]
        lines.append("")
        # Show selection summary
        selected_count = sum(1 for _, (_, sel) in self.odt_locales.items() if sel)
        lines.append(f"Selected: {selected_count} language(s)")
        lines.append("")
        pane = self.panes["odt_locales"]
        entries = self._ensure_pane_lines(pane)
        active_filter = self._get_pane_filter(pane.name)
        if active_filter:
            lines.append(f"Filter: {active_filter}")
        if entries:
            for index, (_, label) in enumerate(entries):
                cursor = "➤" if pane.cursor == index else " "
                lines.append(f"{cursor} {label}")
        else:
            lines.append("No matching locales.")
        lines.append("")
        lines.append("─" * min(width - 2, 50))
        lines.append("Toggle languages with Space. Use / to filter.")
        lines.append("Multiple languages can be selected for Office install.")
        return [line[:width] for line in lines]
        return [line[:width] for line in lines]

    def _render_odt_repair_pane(self, width: int) -> list[str]:
        """Render the ODT repair presets pane."""
        lines = ["ODT Repair Presets:"]
        lines.append("")
        pane = self.panes["odt_repair"]
        entries = self._ensure_pane_lines(pane)
        active_filter = self._get_pane_filter(pane.name)
        if active_filter:
            lines.append(f"Filter: {active_filter}")
        if entries:
            for index, (_, label) in enumerate(entries):
                cursor = "➤" if pane.cursor == index else " "
                lines.append(f"{cursor} {label}")
        else:
            lines.append("No matching presets.")
        lines.append("")
        lines.append("─" * min(width - 2, 50))
        lines.append("Select a preset with Space, Enter/F10 to execute.")
        lines.append("")
        lines.append("• Quick Repair: Fast local repair (no internet)")
        lines.append("• Full Repair: Complete online repair (needs internet)")
        lines.append("• Full Removal: Complete uninstall of Office")
        return [line[:width] for line in lines]

    def _render_run_pane(self, width: int) -> list[str]:
        """Render the execution progress pane."""
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
        """Render the logs pane."""
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
        """Render the settings pane."""
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

    def _render_c2r_remove_pane(self, width: int) -> list[str]:
        """Render the C2R uninstall pane."""
        lines = ["Click-to-Run Uninstall:"]
        lines.append("")
        lines.append("Removes Office Click-to-Run installations using the")
        lines.append("native OfficeClickToRun.exe uninstaller.")
        lines.append("")
        lines.append("─" * 50)
        lines.append("This action will:")
        lines.append("  • Detect installed C2R Office products")
        lines.append("  • Invoke the built-in uninstaller for each")
        lines.append("  • Clean up remaining registry entries")
        lines.append("")
        lines.append("Press Enter or F10 to execute.")
        return [line[:width] for line in lines]

    def _render_offscrub_pane(self, width: int) -> list[str]:
        """Render the OffScrub scripts pane."""
        lines = ["OffScrub Scripts:"]
        lines.append("")
        lines.append("Legacy Microsoft Office removal scripts supporting:")
        lines.append("  • Office 2003, 2007, 2010 (OffScrub03/07/10.vbs)")
        lines.append("  • Office 2013, 2016 MSI (OffScrub_O15/O16msi.vbs)")
        lines.append("  • Office 365/2019/2021 C2R (OffScrubC2R.vbs)")
        lines.append("")
        lines.append("─" * 50)
        lines.append("This action will:")
        lines.append("  • Detect Office versions present")
        lines.append("  • Run appropriate OffScrub script for each")
        lines.append("  • Perform deep registry and file cleanup")
        lines.append("")
        lines.append("Press Enter or F10 to execute.")
        return [line[:width] for line in lines]

    def _render_licensing_pane(self, width: int) -> list[str]:
        """Render the licensing cleanup pane."""
        lines = ["Licensing Cleanup:"]
        lines.append("")
        lines.append("Removes Office product keys and activation state:")
        lines.append("  • Software Protection Platform (SPP) tokens")
        lines.append("  • OSPP (Office Software Protection) entries")
        lines.append("  • Registry-based license information")
        lines.append("  • vNext identity tokens")
        lines.append("")
        lines.append("─" * 50)
        lines.append("This action will:")
        lines.append("  • Remove installed product keys")
        lines.append("  • Clear activation cache")
        lines.append("  • Reset licensing registry keys")
        lines.append("")
        lines.append("Press Enter or F10 to execute.")
        return [line[:width] for line in lines]

    def _render_license_status_pane(self, width: int) -> list[str]:
        """Render the license status pane."""
        lines = ["Licensing Status:"]
        lines.append("")
        lines.append("Query and display current Office licensing state.")
        lines.append("")
        lines.append("─" * 50)
        lines.append("This will show:")
        lines.append("  • Installed Office products")
        lines.append("  • License status for each product")
        lines.append("  • Activation state")
        lines.append("  • OSPP.VBS availability")
        lines.append("")
        lines.append("Press Enter or F10 to query status.")
        return [line[:width] for line in lines]

    # Abstract methods that must be provided by the main class
    def _ensure_pane_lines(self, pane: PaneContext) -> list[tuple[str, str]]:
        """Ensure pane lines are populated. Must be implemented by main class."""
        raise NotImplementedError  # pragma: no cover

    def _get_pane_filter(self, pane_name: str) -> str:
        """Get filter for a pane. Must be implemented by main class."""
        raise NotImplementedError  # pragma: no cover
