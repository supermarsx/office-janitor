"""!
@brief Tests for the text-based user interface engine.
"""

from __future__ import annotations

from collections import deque
from io import StringIO
from types import SimpleNamespace

from src.office_janitor import tui, tui_helpers


def test_enable_windows_ansi_returns_bool(monkeypatch):
    """Test that _enable_windows_ansi returns a boolean without crashing."""
    # Reset the global state
    monkeypatch.setattr(tui_helpers, "_ansi_enabled", False)
    result = tui_helpers._enable_windows_ansi()
    assert isinstance(result, bool)


def test_supports_ansi_tries_enable_on_windows(monkeypatch):
    """Test that supports_ansi attempts to enable ANSI on Windows."""
    import os

    # Simulate Windows without env vars
    monkeypatch.setattr(os, "name", "nt")
    monkeypatch.delenv("WT_SESSION", raising=False)
    monkeypatch.delenv("ANSICON", raising=False)
    monkeypatch.delenv("TERM_PROGRAM", raising=False)
    monkeypatch.delenv("ConEmuANSI", raising=False)

    # Mock a non-tty stream - should return False
    fake_stream = StringIO()
    assert tui_helpers.supports_ansi(fake_stream) is False

    # Reset state and mock successful enable
    monkeypatch.setattr(tui_helpers, "_ansi_enabled", False)
    enable_called = []

    def mock_enable():
        enable_called.append(True)
        return True

    monkeypatch.setattr(tui_helpers, "_enable_windows_ansi", mock_enable)

    # Create a fake tty stream
    class FakeTTY:
        def isatty(self):
            return True

    result = tui_helpers.supports_ansi(FakeTTY())
    assert enable_called == [True]
    assert result is True


def _make_app_state():
    calls: dict[str, object] = {
        "detect": 0,
        "run": 0,
        "planner_overrides": None,
        "executor_overrides": None,
    }

    def detector() -> dict[str, list[str]]:
        calls["detect"] = int(calls["detect"]) + 1
        return {"msi": ["office"]}

    def planner(inventory, overrides):
        calls["planner_overrides"] = dict(overrides or {}) if overrides is not None else None
        return [{"step": "noop"}]

    def executor(plan, overrides):
        calls["run"] = int(calls["run"]) + 1
        calls["executor_overrides"] = dict(overrides or {}) if overrides is not None else None

    state = {
        "detector": detector,
        "planner": planner,
        "executor": executor,
        "event_queue": deque(),
        "confirm": lambda **kwargs: True,
        "input": lambda prompt="": "",  # Mock input to avoid blocking in tests
        "args": SimpleNamespace(
            tui=True,
            quiet=False,
            json=False,
            no_color=False,
            tui_refresh=50,
            tui_compact=False,
            dry_run=False,
            no_restore_point=False,
            no_license=False,
            keep_templates=False,
        ),
    }
    return state, calls


def _select_mode(interface, mode: str) -> None:
    """Helper to select a mode in the TUI.

    Args:
        interface: The TUI interface
        mode: One of 'install', 'repair', 'remove', 'diagnose'
    """
    mode_ids = [m[0] for m in interface.mode_options]
    target_index = mode_ids.index(mode)
    # Navigate to the mode
    while interface.mode_index != target_index:
        if interface.mode_index < target_index:
            interface._handle_key("down")
        else:
            interface._handle_key("up")
    # Select it
    interface._handle_key("enter")


def test_decode_key_windows_arrow_sequences():
    assert tui._decode_key("\x00H") == "up"
    assert tui._decode_key("\x00P") == "down"
    assert tui._decode_key("\x00K") == "left"
    assert tui._decode_key("\x00M") == "right"
    assert tui._decode_key("\xe0H") == "up"
    assert tui._decode_key("\xe0P") == "down"
    assert tui._decode_key("\xe0K") == "left"
    assert tui._decode_key("\xe0M") == "right"


def test_mode_selection_initial_state(monkeypatch):
    """Test TUI starts in mode selection screen."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    assert interface.current_mode is None
    assert interface.mode_index == 0
    assert interface.active_tab == "mode_select"
    assert len(interface.mode_options) == 9  # install, repair, remove, diagnose, odt, offscrub, c2r, license, config
    assert interface.navigation == []  # Navigation is empty until mode selected


def test_mode_selection_navigation(monkeypatch):
    """Test navigating through mode selection options."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    # Initially at index 0
    assert interface.mode_index == 0

    # Navigate down
    interface._handle_key("down")
    assert interface.mode_index == 1

    interface._handle_key("down")
    assert interface.mode_index == 2

    # Navigate up
    interface._handle_key("up")
    assert interface.mode_index == 1

    # Navigate back up to wrap (9 modes total, index 0-8)
    interface._handle_key("up")
    interface._handle_key("up")  # Should wrap to bottom (index 8)
    assert interface.mode_index == 8


def test_mode_selection_right_arrow_enters_mode(monkeypatch):
    """Test right arrow key enters the selected mode."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    # Select install mode using right arrow
    interface._handle_key("right")
    assert interface.current_mode == "install"
    assert len(interface.navigation) > 0


def test_left_arrow_returns_to_mode_selection(monkeypatch):
    """Test left arrow key returns to mode selection from a mode."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    # Enter remove mode
    _select_mode(interface, "remove")
    assert interface.current_mode == "remove"

    # Press left arrow to return to mode selection
    interface._handle_key("left")
    assert interface.current_mode is None
    assert interface.active_tab == "mode_select"


def test_navigation_right_arrow_activates_item(monkeypatch):
    """Test right arrow key activates navigation items within a mode."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    # Enter install mode
    _select_mode(interface, "install")
    assert interface.current_mode == "install"

    # Navigate to odt_presets and activate with right arrow
    interface.focus_area = "nav"
    interface.nav_index = 0  # odt_presets is first
    prev_tab = interface.active_tab
    interface._handle_key("right")
    # Should activate the navigation item (odt_presets has an action)
    assert interface.active_tab == "odt_presets"


def test_mode_selection_enters_mode(monkeypatch):
    """Test selecting a mode enters that mode's navigation."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    # Select install mode
    _select_mode(interface, "install")
    assert interface.current_mode == "install"
    assert len(interface.navigation) > 0
    assert any(item.name == "odt_presets" for item in interface.navigation)

    # Return to mode selection
    interface._return_to_mode_selection()
    assert interface.current_mode is None
    assert interface.active_tab == "mode_select"


def test_escape_returns_to_mode_selection(monkeypatch):
    """Test Escape key returns to mode selection from a mode."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    # Enter remove mode
    _select_mode(interface, "remove")
    assert interface.current_mode == "remove"

    # Press escape to return to mode selection
    interface._handle_key("escape")
    assert interface.current_mode is None
    assert interface.active_tab == "mode_select"


def test_mode_specific_navigation(monkeypatch):
    """Test each mode has its own specific navigation items."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    # Install mode
    _select_mode(interface, "install")
    install_names = [item.name for item in interface.navigation]
    assert "odt_presets" in install_names
    assert "odt_locales" in install_names
    assert "back" in install_names

    # Repair mode
    interface._return_to_mode_selection()
    _select_mode(interface, "repair")
    repair_names = [item.name for item in interface.navigation]
    assert "repair_quick" in repair_names
    assert "repair_full" in repair_names
    assert "repair_odt" in repair_names

    # Remove mode
    interface._return_to_mode_selection()
    _select_mode(interface, "remove")
    remove_names = [item.name for item in interface.navigation]
    assert "detect" in remove_names
    assert "auto" in remove_names
    assert "targeted" in remove_names

    # Diagnose mode
    interface._return_to_mode_selection()
    _select_mode(interface, "diagnose")
    diagnose_names = [item.name for item in interface.navigation]
    assert "detect" in diagnose_names
    assert "diagnostics" in diagnose_names
    assert "plan" in diagnose_names


def test_navigation_state_changes(monkeypatch):
    state, calls = _make_app_state()

    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    # First select "diagnose" mode to get access to detection and plan
    _select_mode(interface, "diagnose")

    # Navigate to the detect item (first item)
    interface.nav_index = 0
    interface.active_tab = interface.navigation[0].name
    interface._handle_key("enter")
    assert calls["detect"] == 1

    # Navigate to plan by finding it
    for idx, item in enumerate(interface.navigation):
        if item.name == "plan":
            interface.nav_index = idx
            interface.active_tab = "plan"
            break

    interface._handle_key("enter")
    assert interface.focus_area == "content"

    plan_cursor_before = interface.panes["plan"].cursor
    interface._handle_key("space")
    assert interface.plan_overrides["include_visio"] is True
    interface._handle_key("down")
    assert interface.panes["plan"].cursor == plan_cursor_before + 1


def test_plan_filter_limits_visible_options(monkeypatch):
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    pane = interface.panes["plan"]
    interface._set_pane_filter("plan", "visio")
    interface._ensure_pane_lines(pane)

    assert pane.lines == ["include_visio"]
    interface._toggle_plan_option(pane.cursor)
    assert interface.plan_overrides["include_visio"] is True

    interface._set_pane_filter("plan", "")


def test_event_queue_updates_state(monkeypatch):
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.event_queue.append(
        {
            "message": "Progress update",
            "data": {
                "status": "Running",
                "log_line": "step completed",
                "inventory": {"c2r": []},
            },
        }
    )

    assert interface._drain_events() is True
    assert interface.progress_message == "Running"
    assert "step completed" in interface.log_lines
    assert interface.last_inventory == {"c2r": []}


def test_fallback_to_cli(monkeypatch, capsys):
    state, _ = _make_app_state()

    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: False)

    interface = tui.OfficeJanitorTUI(state)
    interface.run()

    captured = capsys.readouterr()
    assert "ANSI terminal support" in captured.out or "ANSI" in captured.out


def test_format_inventory_flattens_entries():
    inventory = {
        "msi": [
            {"product": "Office 2019", "arch": "x64"},
            {"product": "Visio", "arch": "x86"},
        ],
        "summary": {"msi": 2},
    }

    lines = tui._format_inventory(inventory)

    assert "msi: product=Office 2019, arch=x64" in lines
    assert "summary.msi: 2" in lines


def test_settings_and_plan_overrides_propagate(monkeypatch):
    state, calls = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.settings_overrides["dry_run"] = True
    interface.settings_overrides["license_cleanup"] = False
    interface.plan_overrides["include_visio"] = True

    interface._handle_plan()
    assert calls["planner_overrides"]["dry_run"] is True
    assert calls["planner_overrides"]["no_license"] is True
    assert calls["planner_overrides"]["include"] == "visio"

    interface._handle_run()
    assert calls["executor_overrides"]["dry_run"] is True
    assert calls["executor_overrides"]["no_license"] is True
    assert calls["executor_overrides"]["include"] == "visio"
    assert calls["executor_overrides"]["confirmed"] is True


def test_targeted_scrub_passes_selected_targets(monkeypatch):
    state, calls = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.target_overrides["2016"] = True
    interface._handle_targeted(execute=True)

    assert calls["planner_overrides"]["target"] == "2016"
    assert calls["planner_overrides"]["mode"] == "target:2016"
    assert calls["executor_overrides"]["target"] == "2016"
    assert calls["executor_overrides"]["mode"] == "target:2016"
    assert calls["executor_overrides"]["confirmed"] is True
    assert calls["run"] == 1


def test_executor_cancellation_updates_status(monkeypatch):
    state, calls = _make_app_state()

    def cancelling_executor(plan, overrides):
        calls["run"] = int(calls["run"]) + 1
        calls["executor_overrides"] = dict(overrides or {}) if overrides is not None else None
        return False

    state["executor"] = cancelling_executor

    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    monkeypatch.setattr(tui, "_spinner", lambda duration, message: None)

    interface = tui.OfficeJanitorTUI(state)
    interface._handle_auto_all()

    assert calls["run"] == 1
    assert interface.progress_message == "Auto Scrub cancelled"
    assert interface.status_lines[-1] == "Auto Scrub cancelled"


def test_confirmation_decline_skips_executor(monkeypatch):
    state, calls = _make_app_state()

    def declining_confirm(**kwargs):
        return False

    state["confirm"] = declining_confirm

    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    monkeypatch.setattr(tui, "_spinner", lambda duration, message: None)

    interface = tui.OfficeJanitorTUI(state)
    interface._handle_auto_all()

    assert calls["run"] == 0
    assert interface.progress_message == "Auto Scrub cancelled"
    assert interface.status_lines[-1] == "Auto Scrub cancelled"


# ---------------------------------------------------------------------------
# ODT Install/Repair Tests
# ---------------------------------------------------------------------------


def test_odt_install_presets_initialized(monkeypatch):
    """Test ODT install presets are properly initialized."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    assert hasattr(interface, "odt_install_presets")
    assert len(interface.odt_install_presets) > 0
    assert "proplus-x64" in interface.odt_install_presets
    assert "office2021-x64" in interface.odt_install_presets
    # Check preset structure: (description, selected)
    desc, selected = interface.odt_install_presets["proplus-x64"]
    assert isinstance(desc, str)
    assert isinstance(selected, bool)
    assert not selected  # Initially not selected


def test_odt_repair_presets_initialized(monkeypatch):
    """Test ODT repair presets are properly initialized."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    assert hasattr(interface, "odt_repair_presets")
    assert len(interface.odt_repair_presets) > 0
    assert "quick-repair" in interface.odt_repair_presets
    assert "full-repair" in interface.odt_repair_presets
    assert "full-removal" in interface.odt_repair_presets


def test_odt_install_preset_selection(monkeypatch):
    """Test selecting an ODT install preset toggles radio-button style."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    # Navigate to ODT install pane
    interface.active_tab = "odt_install"
    pane = interface.panes["odt_install"]
    interface._ensure_pane_lines(pane)

    # Select first preset
    interface._select_odt_install_preset(0)
    first_key = list(interface.odt_install_presets.keys())[0]
    desc, selected = interface.odt_install_presets[first_key]
    assert selected is True
    assert interface.selected_odt_preset == first_key

    # Select second preset - should deselect first
    interface._select_odt_install_preset(1)
    second_key = list(interface.odt_install_presets.keys())[1]
    desc, selected = interface.odt_install_presets[first_key]
    assert selected is False  # First should be deselected
    desc, selected = interface.odt_install_presets[second_key]
    assert selected is True  # Second should be selected


def test_odt_repair_preset_selection(monkeypatch):
    """Test selecting an ODT repair preset."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_repair"
    pane = interface.panes["odt_repair"]
    interface._ensure_pane_lines(pane)

    # Select quick-repair
    quick_repair_idx = list(interface.odt_repair_presets.keys()).index("quick-repair")
    interface._select_odt_repair_preset(quick_repair_idx)

    desc, selected = interface.odt_repair_presets["quick-repair"]
    assert selected is True


def test_odt_install_requires_selection(monkeypatch):
    """Test ODT install requires a preset to be selected."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    # Try to execute without selection
    interface._handle_odt_install(execute=True)
    assert "Select an ODT installation preset" in interface.status_lines[-1]


def test_odt_repair_requires_selection(monkeypatch):
    """Test ODT repair requires a preset to be selected."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    # Try to execute without selection
    interface._handle_odt_repair(execute=True)
    assert "Select an ODT repair preset" in interface.status_lines[-1]


def test_navigation_includes_odt_items(monkeypatch):
    """Test navigation includes ODT Install and Repair items in install/repair modes."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    # Check install mode has odt_presets
    _select_mode(interface, "install")
    install_nav_names = [item.name for item in interface.navigation]
    assert "odt_presets" in install_nav_names

    # Check repair mode has repair_odt
    interface._return_to_mode_selection()
    _select_mode(interface, "repair")
    repair_nav_names = [item.name for item in interface.navigation]
    assert "repair_odt" in repair_nav_names


def test_odt_pane_lines_populated(monkeypatch):
    """Test ODT pane lines are populated correctly."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    # Test odt_install pane
    interface.active_tab = "odt_install"
    pane = interface.panes["odt_install"]
    entries = interface._ensure_pane_lines(pane)
    assert len(entries) == len(interface.odt_install_presets)

    # Test odt_repair pane
    interface.active_tab = "odt_repair"
    pane = interface.panes["odt_repair"]
    entries = interface._ensure_pane_lines(pane)
    assert len(entries) == len(interface.odt_repair_presets)


def test_odt_key_handling_space(monkeypatch):
    """Test space key selects ODT presets."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    # Navigate to odt_install and press space
    interface.active_tab = "odt_install"
    interface.focus_area = "content"
    interface._handle_content_key("space")

    # First preset should be selected
    first_key = list(interface.odt_install_presets.keys())[0]
    desc, selected = interface.odt_install_presets[first_key]
    assert selected is True


# ---------------------------------------------------------------------------
# ODT Locale Selection Tests
# ---------------------------------------------------------------------------


def test_odt_locales_initialized(monkeypatch):
    """Test ODT locales dictionary is properly initialized."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    assert hasattr(interface, "odt_locales")
    assert len(interface.odt_locales) > 0
    assert "en-us" in interface.odt_locales
    assert "de-de" in interface.odt_locales
    assert "fr-fr" in interface.odt_locales
    # Check en-us is selected by default
    desc, selected = interface.odt_locales["en-us"]
    assert selected is True


def test_odt_locale_toggle(monkeypatch):
    """Test toggling ODT locale selections (checkbox style)."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_locales"
    pane = interface.panes["odt_locales"]
    interface._ensure_pane_lines(pane)

    # Get index of de-de
    keys = list(interface.odt_locales.keys())
    de_index = keys.index("de-de")

    # Toggle de-de on
    interface._toggle_odt_locale(de_index)
    desc, selected = interface.odt_locales["de-de"]
    assert selected is True

    # Toggle de-de off
    interface._toggle_odt_locale(de_index)
    desc, selected = interface.odt_locales["de-de"]
    assert selected is False

    # en-us should still be selected (multiple allowed)
    desc, selected = interface.odt_locales["en-us"]
    assert selected is True


def test_odt_locale_multiple_selection(monkeypatch):
    """Test multiple locales can be selected simultaneously."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_locales"
    pane = interface.panes["odt_locales"]
    interface._ensure_pane_lines(pane)

    keys = list(interface.odt_locales.keys())

    # Select German and French in addition to default English
    interface._toggle_odt_locale(keys.index("de-de"))
    interface._toggle_odt_locale(keys.index("fr-fr"))

    selected = interface._get_selected_odt_locales()
    assert "en-us" in selected
    assert "de-de" in selected
    assert "fr-fr" in selected
    assert len(selected) == 3


def test_odt_install_requires_locale(monkeypatch):
    """Test ODT install requires at least one locale selected."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    # Select a preset
    interface._select_odt_install_preset(0)

    # Deselect all locales
    for key in interface.odt_locales:
        desc, _ = interface.odt_locales[key]
        interface.odt_locales[key] = (desc, False)

    # Try to execute
    interface._handle_odt_install(execute=True)
    assert "Select at least one language" in interface.status_lines[-1]


def test_navigation_includes_odt_locales(monkeypatch):
    """Test navigation includes ODT Locales item in install mode."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    # Check install mode has odt_locales
    _select_mode(interface, "install")
    nav_names = [item.name for item in interface.navigation]
    assert "odt_locales" in nav_names


def test_odt_locales_pane_lines_populated(monkeypatch):
    """Test ODT locales pane lines are populated correctly."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_locales"
    pane = interface.panes["odt_locales"]
    entries = interface._ensure_pane_lines(pane)
    assert len(entries) == len(interface.odt_locales)


def test_odt_locales_filter(monkeypatch):
    """Test filtering locales by name."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_locales"
    pane = interface.panes["odt_locales"]

    # Filter for "German"
    interface._set_pane_filter("odt_locales", "German")
    entries = interface._ensure_pane_lines(pane)
    assert len(entries) == 1
    assert entries[0][0] == "de-de"

    # Clear filter
    interface._set_pane_filter("odt_locales", "")
    entries = interface._ensure_pane_lines(pane)
    assert len(entries) == len(interface.odt_locales)


# ---------------------------------------------------------------------------
# ODT Rendering Tests
# ---------------------------------------------------------------------------


def test_render_odt_install_pane(monkeypatch):
    """Test ODT install pane rendering output."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_install"
    lines = interface._render_odt_install_pane(80)

    assert lines[0] == "ODT Installation Presets:"
    assert any("Space" in line for line in lines)
    assert any("Languages:" in line for line in lines)


def test_render_odt_locales_pane(monkeypatch):
    """Test ODT locales pane rendering output."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_locales"
    lines = interface._render_odt_locales_pane(80)

    assert lines[0] == "ODT Language Selection:"
    assert any("Selected:" in line for line in lines)
    assert any("Space toggle" in line for line in lines)  # New help text
    assert any("PgUp/PgDn" in line or "scroll" in line for line in lines)  # Scroll hint


def test_render_odt_repair_pane(monkeypatch):
    """Test ODT repair pane rendering output."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_repair"
    lines = interface._render_odt_repair_pane(80)

    assert lines[0] == "ODT Repair Presets:"
    assert any("Quick Repair" in line for line in lines)
    assert any("Full Repair" in line for line in lines)
    assert any("Full Removal" in line for line in lines)


def test_odt_locales_scrolling(monkeypatch):
    """Test ODT locales pane scrolling with many items."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_locales"
    pane = interface.panes["odt_locales"]

    # There are many locales (40+), scrolling should work
    total_locales = len(interface.odt_locales)
    assert total_locales > 12  # More than max_visible

    # Initially at top
    pane.cursor = 0
    pane.scroll_offset = 0
    lines = interface._render_odt_locales_pane(80)
    assert any("more below" in line for line in lines)  # Should show scroll indicator

    # Move cursor down past visible window
    pane.cursor = 20
    lines = interface._render_odt_locales_pane(80)
    assert any("more above" in line for line in lines)  # Should show top scroll indicator
    assert any("more below" in line for line in lines)  # Should show bottom scroll indicator

    # Move to end
    pane.cursor = total_locales - 1
    lines = interface._render_odt_locales_pane(80)
    assert any("more above" in line for line in lines)  # Should show scroll indicator at top


def test_render_odt_install_shows_locales_summary(monkeypatch):
    """Test ODT install pane shows selected locales summary."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    # Select multiple locales
    for key in ["de-de", "fr-fr", "es-es"]:
        desc, _ = interface.odt_locales[key]
        interface.odt_locales[key] = (desc, True)

    interface.active_tab = "odt_install"
    lines = interface._render_odt_install_pane(80)

    # Should show en-us (default) + our 3 selections = 4 total
    languages_line = [line for line in lines if "Languages:" in line][0]
    assert "en-us" in languages_line


def test_render_odt_locales_shows_count(monkeypatch):
    """Test ODT locales pane shows correct selection count."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    # Default has en-us selected
    interface.active_tab = "odt_locales"
    lines = interface._render_odt_locales_pane(80)

    selected_line = [line for line in lines if "Selected:" in line][0]
    assert "1 language" in selected_line

    # Add more selections
    for key in ["de-de", "fr-fr"]:
        desc, _ = interface.odt_locales[key]
        interface.odt_locales[key] = (desc, True)

    lines = interface._render_odt_locales_pane(80)
    selected_line = [line for line in lines if "Selected:" in line][0]
    assert "3 language" in selected_line


# ---------------------------------------------------------------------------
# ODT Key Handling Tests
# ---------------------------------------------------------------------------


def test_odt_key_handling_f10(monkeypatch):
    """Test F10 key triggers ODT execution."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_install"
    interface.focus_area = "content"

    # Without selection, should show error in status
    interface._handle_content_key("f10")
    assert "Select an ODT installation preset" in interface.status_lines[-1]


def test_odt_locales_key_handling_space(monkeypatch):
    """Test space key toggles locale selection."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_locales"
    interface.focus_area = "content"

    # Move cursor to second item (de-de assuming sorted after en-us)
    interface._handle_content_key("down")
    interface._handle_content_key("space")

    # Check that at least two locales are now selected
    selected_count = sum(1 for _, (_, sel) in interface.odt_locales.items() if sel)
    assert selected_count >= 2


def test_odt_enter_without_selection(monkeypatch):
    """Test Enter on ODT pane without selection shows error."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_repair"
    interface.focus_area = "content"

    interface._handle_content_key("enter")
    assert "Select an ODT repair preset" in interface.status_lines[-1]


# ---------------------------------------------------------------------------
# ODT Prepare Handlers Tests
# ---------------------------------------------------------------------------


def test_prepare_odt_install(monkeypatch):
    """Test _prepare_odt_install sets appropriate state."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface._prepare_odt_install()

    assert "Select ODT installation preset" in interface.progress_message
    assert any("ODT Install" in line for line in interface.status_lines)


def test_prepare_odt_locales(monkeypatch):
    """Test _prepare_odt_locales sets appropriate state."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface._prepare_odt_locales()

    assert "Select Office languages" in interface.progress_message
    assert any("ODT Locales" in line for line in interface.status_lines)


def test_prepare_odt_repair(monkeypatch):
    """Test _prepare_odt_repair sets appropriate state."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface._prepare_odt_repair()

    assert "Select ODT repair preset" in interface.progress_message
    assert any("ODT Repair" in line for line in interface.status_lines)


# ---------------------------------------------------------------------------
# ODT Selection Boundary Tests
# ---------------------------------------------------------------------------


def test_odt_install_preset_selection_out_of_bounds(monkeypatch):
    """Test preset selection handles out of bounds indices gracefully."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_install"
    pane = interface.panes["odt_install"]
    interface._ensure_pane_lines(pane)

    # Try negative index - should clamp to 0
    interface._select_odt_install_preset(-1)
    first_key = list(interface.odt_install_presets.keys())[0]
    desc, selected = interface.odt_install_presets[first_key]
    assert selected is True

    # Try very large index - should clamp to last
    interface._select_odt_install_preset(9999)
    last_key = list(interface.odt_install_presets.keys())[-1]
    desc, selected = interface.odt_install_presets[last_key]
    assert selected is True


def test_odt_locale_toggle_out_of_bounds(monkeypatch):
    """Test locale toggle handles out of bounds indices gracefully."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_locales"
    pane = interface.panes["odt_locales"]
    interface._ensure_pane_lines(pane)

    # Toggle with negative index - should clamp to 0
    first_key = list(interface.odt_locales.keys())[0]
    desc, before = interface.odt_locales[first_key]

    interface._toggle_odt_locale(-1)
    desc, after = interface.odt_locales[first_key]
    assert after != before


# ---------------------------------------------------------------------------
# ODT Filter Tests
# ---------------------------------------------------------------------------


def test_odt_install_filter(monkeypatch):
    """Test filtering ODT install presets."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_install"
    pane = interface.panes["odt_install"]

    # Filter for "2021"
    interface._set_pane_filter("odt_install", "2021")
    entries = interface._ensure_pane_lines(pane)

    # Should match office2021-x64
    assert len(entries) >= 1
    assert all("2021" in entry[1].lower() for entry in entries)

    # Clear filter
    interface._set_pane_filter("odt_install", "")
    entries = interface._ensure_pane_lines(pane)
    assert len(entries) == len(interface.odt_install_presets)


def test_odt_repair_filter(monkeypatch):
    """Test filtering ODT repair presets."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_repair"
    pane = interface.panes["odt_repair"]

    # Filter for "repair" - should match quick-repair and full-repair
    interface._set_pane_filter("odt_repair", "repair")
    entries = interface._ensure_pane_lines(pane)

    # Should match quick-repair and full-repair (not full-removal)
    assert len(entries) == 2
    assert all("repair" in entry[0].lower() for entry in entries)


# ---------------------------------------------------------------------------
# ODT Get Selected Locales Tests
# ---------------------------------------------------------------------------


def test_get_selected_odt_locales_default(monkeypatch):
    """Test _get_selected_odt_locales returns default selection."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    selected = interface._get_selected_odt_locales()
    assert "en-us" in selected
    assert len(selected) == 1


def test_get_selected_odt_locales_all_deselected(monkeypatch):
    """Test _get_selected_odt_locales when all are deselected."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    # Deselect all
    for key in interface.odt_locales:
        desc, _ = interface.odt_locales[key]
        interface.odt_locales[key] = (desc, False)

    selected = interface._get_selected_odt_locales()
    assert len(selected) == 0


# ---------------------------------------------------------------------------
# ODT Content Dispatch Tests
# ---------------------------------------------------------------------------


def test_render_content_dispatches_odt_install(monkeypatch):
    """Test _render_content dispatches to ODT install pane."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_install"
    lines = interface._render_content(80)

    assert lines[0] == "ODT Installation Presets:"


def test_render_content_dispatches_odt_locales(monkeypatch):
    """Test _render_content dispatches to ODT locales pane."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_locales"
    lines = interface._render_content(80)

    assert lines[0] == "ODT Language Selection:"


def test_render_content_dispatches_odt_repair(monkeypatch):
    """Test _render_content dispatches to ODT repair pane."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_repair"
    lines = interface._render_content(80)

    assert lines[0] == "ODT Repair Presets:"


# ---------------------------------------------------------------------------
# ODT Cursor Movement Tests
# ---------------------------------------------------------------------------


def test_odt_cursor_up_down(monkeypatch):
    """Test cursor movement in ODT panes."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_install"
    interface.focus_area = "content"
    pane = interface.panes["odt_install"]

    # Initial cursor should be 0
    assert pane.cursor == 0

    # Move down
    interface._handle_content_key("down")
    assert pane.cursor == 1

    # Move up
    interface._handle_content_key("up")
    assert pane.cursor == 0

    # Move up at top should stay at 0
    interface._handle_content_key("up")
    assert pane.cursor == 0


def test_odt_locales_cursor_bounds(monkeypatch):
    """Test cursor stays within bounds in ODT locales pane."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_locales"
    interface.focus_area = "content"
    pane = interface.panes["odt_locales"]
    interface._ensure_pane_lines(pane)

    # Move cursor to end
    for _ in range(len(interface.odt_locales) + 5):
        interface._handle_content_key("down")

    # Cursor should be at last valid index
    assert pane.cursor == len(pane.lines) - 1


# ---------------------------------------------------------------------------
# ODT Selection Display Tests
# ---------------------------------------------------------------------------


def test_odt_install_displays_selection_state(monkeypatch):
    """Test ODT install pane displays selection state correctly."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_install"
    pane = interface.panes["odt_install"]

    # Select a preset
    interface._select_odt_install_preset(0)
    entries = interface._ensure_pane_lines(pane)

    # First entry should show selected indicator
    first_key, first_label = entries[0]
    assert (
        "●" in first_label
        or "[x]" in first_label.lower()
        or "(selected)" in first_label.lower()
        or "●" in first_label
    )


def test_odt_locales_displays_checkbox_state(monkeypatch):
    """Test ODT locales pane displays checkbox state correctly."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_locales"
    pane = interface.panes["odt_locales"]
    entries = interface._ensure_pane_lines(pane)

    # en-us should be selected by default
    en_us_entry = [e for e in entries if e[0] == "en-us"][0]
    # The label should have a checked indicator
    assert "☑" in en_us_entry[1] or "[x]" in en_us_entry[1].lower() or "●" in en_us_entry[1]


# ---------------------------------------------------------------------------
# ODT Navigation Order Tests
# ---------------------------------------------------------------------------


def test_odt_navigation_order(monkeypatch):
    """Test ODT items appear in expected navigation order within modes."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    # Test install mode navigation order
    _select_mode(interface, "install")
    install_nav_names = [item.name for item in interface.navigation]
    # Install mode should have: odt_presets, odt_locales, odt_custom, run_install, back
    assert install_nav_names.index("odt_presets") < install_nav_names.index("odt_locales")
    assert install_nav_names.index("odt_locales") < install_nav_names.index("odt_custom")
    assert install_nav_names.index("odt_custom") < install_nav_names.index("run_install")

    # Test repair mode navigation order
    interface._return_to_mode_selection()
    _select_mode(interface, "repair")
    repair_nav_names = [item.name for item in interface.navigation]
    # Repair mode should have: repair_quick, repair_full, repair_odt, run_repair, back
    assert repair_nav_names.index("repair_quick") < repair_nav_names.index("repair_full")
    assert repair_nav_names.index("repair_full") < repair_nav_names.index("repair_odt")
    assert repair_nav_names.index("repair_odt") < repair_nav_names.index("run_repair")

    # Test remove mode navigation order
    interface._return_to_mode_selection()
    _select_mode(interface, "remove")
    remove_nav_names = [item.name for item in interface.navigation]
    # Remove mode should have: detect, auto, targeted, cleanup, settings, run_remove, back
    assert remove_nav_names.index("detect") < remove_nav_names.index("auto")
    assert remove_nav_names.index("auto") < remove_nav_names.index("targeted")

    # Test diagnose mode navigation order
    interface._return_to_mode_selection()
    _select_mode(interface, "diagnose")
    diagnose_nav_names = [item.name for item in interface.navigation]
    # Diagnose mode should have: detect, diagnostics, plan, logs, back
    assert diagnose_nav_names.index("detect") < diagnose_nav_names.index("diagnostics")
    assert diagnose_nav_names.index("diagnostics") < diagnose_nav_names.index("plan")


# ---------------------------------------------------------------------------
# ODT Repair Preset Deselection Tests
# ---------------------------------------------------------------------------


def test_odt_repair_radio_button_behavior(monkeypatch):
    """Test ODT repair presets behave as radio buttons."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "odt_repair"
    pane = interface.panes["odt_repair"]
    interface._ensure_pane_lines(pane)

    # Select quick-repair
    keys = list(interface.odt_repair_presets.keys())
    quick_idx = keys.index("quick-repair")
    interface._select_odt_repair_preset(quick_idx)

    # Verify quick-repair is selected
    desc, selected = interface.odt_repair_presets["quick-repair"]
    assert selected is True

    # Select full-repair
    full_idx = keys.index("full-repair")
    interface._select_odt_repair_preset(full_idx)

    # quick-repair should be deselected, full-repair selected
    desc, selected = interface.odt_repair_presets["quick-repair"]
    assert selected is False
    desc, selected = interface.odt_repair_presets["full-repair"]
    assert selected is True


# ---------------------------------------------------------------------------
# Special Action Tests: OffScrub, C2R, Licensing
# ---------------------------------------------------------------------------


def test_remove_mode_has_special_actions(monkeypatch):
    """Test remove mode includes OffScrub, C2R, and Licensing actions."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    _select_mode(interface, "remove")
    nav_names = [item.name for item in interface.navigation]

    assert "c2r_remove" in nav_names
    assert "offscrub" in nav_names
    assert "licensing" in nav_names


def test_diagnose_mode_has_license_status(monkeypatch):
    """Test diagnose mode includes Licensing Status action."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    _select_mode(interface, "diagnose")
    nav_names = [item.name for item in interface.navigation]

    assert "license_status" in nav_names


def test_prepare_c2r_remove(monkeypatch):
    """Test _prepare_c2r_remove sets appropriate messages."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface._prepare_c2r_remove()

    assert "C2R" in interface.progress_message
    assert any("C2R" in line for line in interface.status_lines)


def test_prepare_offscrub(monkeypatch):
    """Test _prepare_offscrub sets appropriate messages."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface._prepare_offscrub()

    assert "OffScrub" in interface.progress_message
    assert any("OffScrub" in line for line in interface.status_lines)


def test_prepare_licensing(monkeypatch):
    """Test _prepare_licensing sets appropriate messages."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface._prepare_licensing()

    assert "Licensing" in interface.progress_message
    assert any("Licensing" in line or "license" in line.lower() for line in interface.status_lines)


def test_new_modes_have_navigation(monkeypatch):
    """Test that all new modes have proper navigation defined."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    new_modes = ["odt", "offscrub", "c2r", "license", "config"]
    for mode in new_modes:
        _select_mode(interface, mode)
        assert interface.current_mode == mode, f"Mode {mode} not set"
        assert len(interface.navigation) > 0, f"Mode {mode} has no navigation"
        assert any(item.name == "back" for item in interface.navigation), f"Mode {mode} missing back item"
        interface._return_to_mode_selection()


def test_odt_mode_navigation_items(monkeypatch):
    """Test ODT mode has expected navigation items."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    _select_mode(interface, "odt")
    nav_names = [item.name for item in interface.navigation]
    assert "odt_presets" in nav_names
    assert "odt_locales" in nav_names
    assert "odt_custom" in nav_names
    assert "odt_export" in nav_names
    assert "back" in nav_names


def test_c2r_mode_navigation_items(monkeypatch):
    """Test C2R mode has expected navigation items."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    _select_mode(interface, "c2r")
    nav_names = [item.name for item in interface.navigation]
    assert "detect" in nav_names
    assert "c2r_remove" in nav_names
    assert "c2r_repair" in nav_names
    assert "c2r_update" in nav_names
    assert "c2r_channel" in nav_names
    assert "back" in nav_names


def test_license_mode_navigation_items(monkeypatch):
    """Test License mode has expected navigation items."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    _select_mode(interface, "license")
    nav_names = [item.name for item in interface.navigation]
    assert "license_status" in nav_names
    assert "license_install" in nav_names
    assert "license_remove" in nav_names
    assert "license_activate" in nav_names
    assert "back" in nav_names


def test_config_mode_navigation_items(monkeypatch):
    """Test Config mode has expected navigation items."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    _select_mode(interface, "config")
    nav_names = [item.name for item in interface.navigation]
    assert "config_view" in nav_names
    assert "config_edit" in nav_names
    assert "config_export" in nav_names
    assert "config_import" in nav_names
    assert "back" in nav_names

def test_render_c2r_remove_pane(monkeypatch):
    """Test C2R remove pane renders correctly."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "c2r_remove"
    lines = interface._render_content(60)

    assert any("Click-to-Run" in line for line in lines)


def test_render_offscrub_pane(monkeypatch):
    """Test OffScrub pane renders correctly."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "offscrub"
    lines = interface._render_content(60)

    assert any("OffScrub" in line for line in lines)


def test_render_licensing_pane(monkeypatch):
    """Test Licensing pane renders correctly."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "licensing"
    lines = interface._render_content(60)

    assert any("Licensing" in line for line in lines)


def test_render_license_status_pane(monkeypatch):
    """Test License status pane renders correctly."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "license_status"
    lines = interface._render_content(60)

    assert any("Licensing Status" in line for line in lines)


def test_render_c2r_channel_pane(monkeypatch):
    """Test C2R channel pane renders with channel options."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    interface.active_tab = "c2r_channel"
    lines = interface._render_content(80)

    assert any("Update Channel" in line for line in lines)
    assert any("Current" in line for line in lines)
    assert any("Monthly" in line or "Semi-Annual" in line for line in lines)


def test_c2r_channel_selection(monkeypatch):
    """Test selecting a C2R channel updates state."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    # Initially no channel selected
    assert interface.selected_c2r_channel is None

    # Select a channel
    interface.active_tab = "c2r_channel"
    pane = interface.panes["c2r_channel"]
    pane.cursor = 1  # monthly
    interface._select_c2r_channel(1)

    # Check that the channel is selected (radio button style)
    selected_count = sum(1 for _, (_, sel) in interface.c2r_channels.items() if sel)
    assert selected_count == 1
    assert interface.selected_c2r_channel is not None
