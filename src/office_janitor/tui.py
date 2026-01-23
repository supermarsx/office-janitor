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

from __future__ import annotations  # noqa: I001

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
)

# Re-export for backward compatibility (tests patch these underscore aliases)
from .tui_helpers import (  # noqa: F401
    clear_screen as _clear_screen,
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
)
from .tui_render import TUIRendererMixin

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
    scroll_offset: int = 0  # For scrolling long lists
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
        self.ansi_supported = _supports_ansi() and not bool(
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

        # Mode selection - user picks a mode first, then sees relevant actions
        # None = mode selection, else install/repair/remove/diagnose
        self.current_mode: str | None = None
        self.mode_options = [
            ("install", "Install Office", "Deploy Office via ODT presets or custom configs"),
            ("repair", "Repair Office", "Fix broken Office installations (quick or full)"),
            ("remove", "Remove Office", "Uninstall Office and clean up residual artifacts"),
            ("diagnose", "Diagnose", "Detect and report Office installations without changes"),
            ("odt", "ODT Builder", "Create custom Office Deployment Tool configurations"),
            ("offscrub", "OffScrub", "Run legacy Microsoft OffScrub removal scripts"),
            ("c2r", "C2R Passthrough", "Direct Click-to-Run OfficeC2RClient commands"),
            ("license", "Licensing", "Manage Office product keys and license activation"),
            ("config", "Config Generator", "Generate configuration files and deployment scripts"),
        ]
        self.mode_index = 0

        # Mode-specific navigation items
        self._install_nav = [
            NavigationItem("odt_install", "Installation Presets", action=self._prepare_odt_install),
            NavigationItem("odt_products", "Select Products", action=self._prepare_odt_products),
            NavigationItem("odt_locales", "Language Selection", action=self._prepare_odt_locales),
            NavigationItem("odt_custom", "Custom Configuration", action=self._prepare_odt_custom),
            NavigationItem("odt_import", "Import ODT File", action=self._prepare_odt_import),
            NavigationItem(
                "run_install", "▶ Run Installation", action=self._handle_odt_install_run
            ),
            NavigationItem("back", "← Back to Modes", action=self._return_to_mode_selection),
        ]
        self._repair_nav = [
            NavigationItem("repair_quick", "Quick Repair", action=self._prepare_repair_quick),
            NavigationItem("repair_full", "Full Online Repair", action=self._prepare_repair_full),
            NavigationItem("repair_odt", "ODT Repair Config", action=self._prepare_odt_repair),
            NavigationItem("run_repair", "▶ Run Repair", action=self._handle_repair_run),
            NavigationItem("back", "← Back to Modes", action=self._return_to_mode_selection),
        ]
        self._remove_nav = [
            NavigationItem("detect", "Detect Inventory", action=self._handle_detect),
            NavigationItem("auto", "Auto Remove All", action=self._handle_auto_all),
            NavigationItem("targeted", "Targeted Remove", action=self._prepare_targeted),
            NavigationItem("scrub_level", "Scrub Level", action=self._prepare_scrub_level),
            NavigationItem("c2r_remove", "C2R Uninstall", action=self._prepare_c2r_remove),
            NavigationItem("offscrub", "OffScrub Scripts", action=self._prepare_offscrub),
            NavigationItem("licensing", "Licensing Cleanup", action=self._prepare_licensing),
            NavigationItem("cleanup", "Cleanup Only", action=self._handle_cleanup_only),
            NavigationItem("settings", "Scrub Settings", action=None),
            NavigationItem("run_remove", "▶ Run Removal", action=self._handle_run),
            NavigationItem("back", "← Back to Modes", action=self._return_to_mode_selection),
        ]
        self._diagnose_nav = [
            NavigationItem("detect", "Detect Inventory", action=self._handle_detect),
            NavigationItem("diagnostics", "Run Diagnostics", action=self._handle_diagnostics),
            NavigationItem(
                "license_status", "Licensing Status", action=self._handle_license_status
            ),
            NavigationItem("plan", "View Plan", action=None),
            NavigationItem("logs", "View Logs", action=self._handle_logs),
            NavigationItem("back", "← Back to Modes", action=self._return_to_mode_selection),
        ]

        # New specialized mode navigation
        self._odt_nav = [
            NavigationItem("odt_install", "Preset Templates", action=self._prepare_odt_install),
            NavigationItem("odt_products", "Select Products", action=self._prepare_odt_products),
            NavigationItem("odt_locales", "Language Selection", action=self._prepare_odt_locales),
            NavigationItem("odt_custom", "Custom XML Editor", action=self._prepare_odt_custom),
            NavigationItem("odt_import", "Import Config File", action=self._prepare_odt_import),
            NavigationItem("odt_export", "Export Config", action=self._handle_odt_export),
            NavigationItem("back", "← Back to Modes", action=self._return_to_mode_selection),
        ]
        self._offscrub_nav = [
            NavigationItem("detect", "Detect Inventory", action=self._handle_detect),
            NavigationItem("offscrub_select", "Select Scripts", action=self._prepare_offscrub),
            NavigationItem("offscrub_run", "▶ Run OffScrub", action=self._handle_offscrub_run),
            NavigationItem("logs", "View Logs", action=self._handle_logs),
            NavigationItem("back", "← Back to Modes", action=self._return_to_mode_selection),
        ]
        self._c2r_nav = [
            NavigationItem("detect", "Detect Inventory", action=self._handle_detect),
            NavigationItem("c2r_remove", "Uninstall Product", action=self._prepare_c2r_remove),
            NavigationItem("c2r_repair", "Repair Product", action=self._prepare_c2r_repair),
            NavigationItem("c2r_update", "Force Update", action=self._handle_c2r_update),
            NavigationItem("c2r_channel", "Change Channel", action=self._prepare_c2r_channel),
            NavigationItem("back", "← Back to Modes", action=self._return_to_mode_selection),
        ]
        self._license_nav = [
            NavigationItem(
                "license_status", "Licensing Status", action=self._handle_license_status
            ),
            NavigationItem(
                "license_install", "Install Product Key", action=self._prepare_license_install
            ),
            NavigationItem("license_remove", "Remove Licenses", action=self._prepare_licensing),
            NavigationItem(
                "license_activate", "Activate Office", action=self._handle_license_activate
            ),
            NavigationItem("back", "← Back to Modes", action=self._return_to_mode_selection),
        ]
        self._config_nav = [
            NavigationItem("config_view", "View Current Config", action=self._handle_config_view),
            NavigationItem("config_edit", "Edit Settings", action=self._prepare_config_edit),
            NavigationItem("config_export", "Export to JSON", action=self._handle_config_export),
            NavigationItem("config_import", "Import Config", action=self._prepare_config_import),
            NavigationItem("back", "← Back to Modes", action=self._return_to_mode_selection),
        ]

        # Current navigation (set when mode is selected)
        self.navigation: list[NavigationItem] = []
        self.focus_area = "nav"
        self.nav_index = 0
        self.active_tab = "mode_select"  # Start at mode selection

        # Create panes for all possible tabs across all modes
        all_tabs = [
            "mode_select",
            "detect",
            "auto",
            "targeted",
            "cleanup",
            "diagnostics",
            "odt_presets",
            "odt_products",
            "odt_locales",
            "odt_custom",
            "odt_import",
            "odt_install",
            "odt_repair",
            "odt_export",
            "repair_quick",
            "repair_full",
            "repair_odt",
            "c2r_remove",
            "c2r_repair",
            "c2r_update",
            "c2r_channel",
            "scrub_level",
            "offscrub",
            "offscrub_select",
            "offscrub_run",
            "licensing",
            "license_status",
            "license_install",
            "license_remove",
            "license_activate",
            "config_view",
            "config_edit",
            "config_export",
            "config_import",
            "plan",
            "run",
            "logs",
            "settings",
            "run_install",
            "run_repair",
            "run_remove",
            "back",
        ]
        self.panes: dict[str, PaneContext] = {name: PaneContext(name) for name in all_tabs}
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
        # ODT Install presets with descriptions
        self.odt_install_presets: dict[str, tuple[str, bool]] = {
            "proplus-x64": ("Microsoft 365 ProPlus (64-bit)", False),
            "proplus-x86": ("Microsoft 365 ProPlus (32-bit)", False),
            "proplus-visio-project": ("ProPlus + Visio + Project", False),
            "business-x64": ("Microsoft 365 Business (64-bit)", False),
            "office2019-x64": ("Office 2019 Professional Plus (64-bit)", False),
            "office2021-x64": ("Office LTSC 2021 (64-bit)", False),
            "office2024-x64": ("Office LTSC 2024 (64-bit)", False),
            "multilang": ("Multi-language Installation", False),
            "shared-computer": ("Shared Computer Activation", False),
            "interactive": ("Interactive Setup (Full UI)", False),
        }
        # ODT Repair presets with descriptions
        self.odt_repair_presets: dict[str, tuple[str, bool]] = {
            "quick-repair": ("Quick Repair (Local)", False),
            "full-repair": ("Full Online Repair", False),
            "full-removal": ("Complete Office Removal", False),
        }
        # Currently selected ODT preset
        self.selected_odt_preset: str | None = None
        # Locale selection for ODT install (code -> (display_name, selected))
        self.odt_locales: dict[str, tuple[str, bool]] = {
            "en-us": ("English (US)", True),  # Default selected
            "en-gb": ("English (UK)", False),
            "de-de": ("German", False),
            "fr-fr": ("French", False),
            "es-es": ("Spanish (Spain)", False),
            "it-it": ("Italian", False),
            "pt-br": ("Portuguese (Brazil)", False),
            "pt-pt": ("Portuguese (Portugal)", False),
            "nl-nl": ("Dutch", False),
            "pl-pl": ("Polish", False),
            "ru-ru": ("Russian", False),
            "ja-jp": ("Japanese", False),
            "zh-cn": ("Chinese (Simplified)", False),
            "zh-tw": ("Chinese (Traditional)", False),
            "ko-kr": ("Korean", False),
            "ar-sa": ("Arabic (Saudi Arabia)", False),
            "he-il": ("Hebrew", False),
            "hi-in": ("Hindi", False),
            "th-th": ("Thai", False),
            "tr-tr": ("Turkish", False),
            "cs-cz": ("Czech", False),
            "da-dk": ("Danish", False),
            "fi-fi": ("Finnish", False),
            "el-gr": ("Greek", False),
            "hu-hu": ("Hungarian", False),
            "nb-no": ("Norwegian (Bokmål)", False),
            "sv-se": ("Swedish", False),
            "uk-ua": ("Ukrainian", False),
            "vi-vn": ("Vietnamese", False),
            "bg-bg": ("Bulgarian", False),
            "hr-hr": ("Croatian", False),
            "et-ee": ("Estonian", False),
            "id-id": ("Indonesian", False),
            "kk-kz": ("Kazakh", False),
            "lv-lv": ("Latvian", False),
            "lt-lt": ("Lithuanian", False),
            "ms-my": ("Malay", False),
            "ro-ro": ("Romanian", False),
            "sr-latn-rs": ("Serbian (Latin)", False),
            "sk-sk": ("Slovak", False),
            "sl-si": ("Slovenian", False),
        }
        # Add locale pane
        self.panes["odt_locales"] = PaneContext("odt_locales")

        # ODT Products selection (product_id -> (display_name, selected))
        self.odt_products: dict[str, tuple[str, bool]] = {
            "O365ProPlusRetail": ("Microsoft 365 Apps for enterprise", False),
            "O365BusinessRetail": ("Microsoft 365 Apps for business", False),
            "ProPlus2024Volume": ("Office LTSC Professional Plus 2024", False),
            "ProPlus2021Volume": ("Office LTSC Professional Plus 2021", False),
            "ProPlus2019Volume": ("Office Professional Plus 2019", False),
            "VisioProRetail": ("Visio Professional (subscription)", False),
            "VisioPro2024Volume": ("Visio Professional 2024", False),
            "VisioPro2021Volume": ("Visio Professional 2021", False),
            "VisioPro2019Volume": ("Visio Professional 2019", False),
            "ProjectProRetail": ("Project Professional (subscription)", False),
            "ProjectPro2024Volume": ("Project Professional 2024", False),
            "ProjectPro2021Volume": ("Project Professional 2021", False),
            "ProjectPro2019Volume": ("Project Professional 2019", False),
            "AccessRetail": ("Access (standalone)", False),
            "Access2024Retail": ("Access 2024", False),
            "Access2021Retail": ("Access 2021", False),
            "Access2019Retail": ("Access 2019", False),
        }
        self.imported_odt_config: str | None = None

        # C2R Update Channel options (channel_id -> (display_name, selected))
        self.c2r_channels: dict[str, tuple[str, bool]] = {
            "current": ("Current Channel (Recommended)", False),
            "monthly": ("Monthly Enterprise Channel", False),
            "semi-annual": ("Semi-Annual Enterprise Channel", False),
            "current-preview": ("Current Channel (Preview)", False),
            "semi-annual-preview": ("Semi-Annual Channel (Preview)", False),
            "beta": ("Beta Channel (Insiders)", False),
        }
        self.selected_c2r_channel: str | None = None

        # Scrub level options (level_id -> (display_name, selected))
        self.scrub_levels: dict[str, tuple[str, bool]] = {
            "minimal": ("Minimal - Remove only installed products", False),
            "standard": ("Standard - Remove products + common artifacts", True),  # Default
            "aggressive": ("Aggressive - Deep cleanup, remove more residual files", False),
            "nuclear": ("Nuclear - Maximum cleanup, may affect shared components", False),
        }
        self.selected_scrub_level = "standard"

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

        # Re-check ANSI support (supports_ansi will try to enable it on Windows)
        if not self.ansi_supported:
            self.ansi_supported = _supports_ansi()

        if not self.ansi_supported:
            self._notify("tui.fallback", "TUI unavailable (ANSI not supported).")
            print(
                "\nTUI requires ANSI terminal support. Your terminal does not appear "
                "to support ANSI escape sequences.\n"
                "Try running from Windows Terminal, VS Code terminal, or a newer "
                "version of PowerShell/cmd.exe (Windows 10 1511+).\n"
            )
            self._wait_for_enter()
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

        # Prompt before closing so user can see final state
        self._wait_for_enter()

    def _wait_for_enter(self) -> None:
        """!
        @brief Wait for the user to press Enter before closing.
        @details Used to prevent the console window from closing immediately.
        Only pauses when running interactively (not in tests or piped input).
        """
        import os
        import sys

        # Don't pause in test environments
        if os.environ.get("PYTEST_CURRENT_TEST"):
            return
        # Don't pause if stdin is not a TTY
        if not sys.stdin.isatty():
            return
        print("\nPress Enter to exit...")
        try:
            input_func = self.app_state.get("input", input)
            if callable(input_func):
                input_func("")
            else:
                input()
        except (EOFError, OSError, KeyboardInterrupt):
            pass  # Non-interactive context

    # -----------------------------------------------------------------------
    # Key handling
    # -----------------------------------------------------------------------

    def _handle_key(self, command: str) -> None:
        """!
        @brief Interpret a normalized key command and update state.
        """

        if command == "quit":
            self.progress_message = "Exiting..."
            self._notify("tui.exit", "User requested exit from TUI.")
            self._running = False
            return

        if command == "escape":
            # Escape returns to mode selection if in a mode
            if self.current_mode is not None:
                self._return_to_mode_selection()
                return
            # Otherwise quit
            self.progress_message = "Exiting..."
            self._notify("tui.exit", "User requested exit from TUI.")
            self._running = False
            return

        if command == "left":
            # Left arrow returns to mode selection if in a mode
            if self.current_mode is not None:
                self._return_to_mode_selection()
                return
            return

        if command == "f1":
            self._show_help()
            return

        # Mode selection screen handling
        if self.current_mode is None:
            self._handle_mode_selection_key(command)
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
            if command in {"enter", "space", "right"}:
                self._activate_nav()
                return
            if command == "f10":
                self._handle_run()
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
            # Larger jump for locale lists, normal for others
            jump = 10 if self.active_tab == "odt_locales" else 5
            pane.cursor = min(max(len(pane.lines) - 1, 0), pane.cursor + jump)
            return
        if command == "page_up":
            jump = 10 if self.active_tab == "odt_locales" else 5
            pane.cursor = max(0, pane.cursor - jump)
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
            elif self.active_tab == "odt_install":
                self._handle_odt_install(execute=True)
            elif self.active_tab == "odt_repair":
                self._handle_odt_repair(execute=True)
            elif self.active_tab == "c2r_remove":
                self._handle_c2r_remove()
            elif self.active_tab == "c2r_channel":
                self._handle_c2r_channel_change()
            elif self.active_tab == "offscrub":
                self._handle_offscrub()
            elif self.active_tab == "licensing":
                self._handle_licensing_cleanup()
            elif self.active_tab == "license_status":
                self._handle_license_status()
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
            elif self.active_tab == "odt_install":
                self._handle_odt_install(execute=True)
            elif self.active_tab == "odt_repair":
                self._handle_odt_repair(execute=True)
            elif self.active_tab == "c2r_remove":
                self._handle_c2r_remove()
            elif self.active_tab == "offscrub":
                self._handle_offscrub()
            elif self.active_tab == "licensing":
                self._handle_licensing_cleanup()
            elif self.active_tab == "license_status":
                self._handle_license_status()
            return
        if command == "space":
            if self.active_tab == "plan":
                self._toggle_plan_option(pane.cursor)
            elif self.active_tab == "settings":
                self._toggle_setting(pane.cursor)
            elif self.active_tab == "targeted":
                self._toggle_target(pane.cursor)
            elif self.active_tab == "odt_install":
                self._select_odt_install_preset(pane.cursor)
            elif self.active_tab == "odt_repair":
                self._select_odt_repair_preset(pane.cursor)
            elif self.active_tab == "odt_locales":
                self._toggle_odt_locale(pane.cursor)
            elif self.active_tab == "odt_products":
                self._toggle_odt_product(pane.cursor)
            elif self.active_tab == "c2r_channel":
                self._select_c2r_channel(pane.cursor)
            elif self.active_tab == "scrub_level":
                self._select_scrub_level(pane.cursor)
            return
        if command == "/":
            self._prompt_filter(pane)
            self._ensure_pane_lines(pane)
            return

    # -----------------------------------------------------------------------
    # Mode selection
    # -----------------------------------------------------------------------

    def _handle_mode_selection_key(self, command: str) -> None:
        """!
        @brief Handle key presses on the mode selection screen.
        """
        if command == "quit":
            self._running = False
            self.progress_message = "Exiting..."
            self._notify("tui.exit", "User requested exit from TUI.")
            return
        if command == "down":
            self.mode_index = (self.mode_index + 1) % len(self.mode_options)
            return
        if command == "up":
            self.mode_index = (self.mode_index - 1) % len(self.mode_options)
            return
        if command in {"enter", "space", "right"}:
            self._select_mode(self.mode_index)
            return

    def _select_mode(self, index: int) -> None:
        """!
        @brief Enter the selected mode and load its navigation.
        """
        mode_id, _, _ = self.mode_options[index]
        self.current_mode = mode_id

        mode_nav_map = {
            "install": self._install_nav,
            "repair": self._repair_nav,
            "remove": self._remove_nav,
            "diagnose": self._diagnose_nav,
            "odt": self._odt_nav,
            "offscrub": self._offscrub_nav,
            "c2r": self._c2r_nav,
            "license": self._license_nav,
            "config": self._config_nav,
        }

        self.navigation = mode_nav_map.get(mode_id, self._remove_nav)
        self.nav_index = 0
        self.active_tab = self.navigation[0].name
        self.focus_area = "nav"
        self._append_status(f"Entered {mode_id.title()} mode")

    def _return_to_mode_selection(self) -> None:
        """!
        @brief Return to the mode selection screen.
        """
        self.current_mode = None
        self.mode_index = 0
        self.active_tab = "mode_select"
        self.focus_area = "nav"
        self._append_status("Returned to mode selection")

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
        elif pane.name == "odt_install":
            entries = [
                (
                    key,
                    f"{'[●]' if selected else '[ ]'} {desc}",
                )
                for key, (desc, selected) in self.odt_install_presets.items()
            ]
        elif pane.name == "odt_repair":
            entries = [
                (
                    key,
                    f"{'[●]' if selected else '[ ]'} {desc}",
                )
                for key, (desc, selected) in self.odt_repair_presets.items()
            ]
        elif pane.name == "odt_locales":
            entries = [
                (
                    key,
                    f"{'[x]' if selected else '[ ]'} {desc} ({key})",
                )
                for key, (desc, selected) in self.odt_locales.items()
            ]
        elif pane.name == "odt_products":
            entries = [
                (
                    key,
                    f"{'[x]' if selected else '[ ]'} {desc}",
                )
                for key, (desc, selected) in self.odt_products.items()
            ]
        elif pane.name == "c2r_channel":
            entries = [
                (
                    key,
                    f"{'[●]' if selected else '[ ]'} {desc}",
                )
                for key, (desc, selected) in self.c2r_channels.items()
            ]
        elif pane.name == "scrub_level":
            entries = [
                (
                    key,
                    f"{'[●]' if selected else '[ ]'} {desc}",
                )
                for key, (desc, selected) in self.scrub_levels.items()
            ]
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
