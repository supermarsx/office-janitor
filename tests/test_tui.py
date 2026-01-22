"""!
@brief Tests for the text-based user interface engine.
"""

from __future__ import annotations

from collections import deque
from types import SimpleNamespace

from src.office_janitor import tui


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


def test_decode_key_windows_arrow_sequences():
    assert tui._decode_key("\x00H") == "up"
    assert tui._decode_key("\x00P") == "down"
    assert tui._decode_key("\x00K") == "left"
    assert tui._decode_key("\x00M") == "right"
    assert tui._decode_key("\xe0H") == "up"
    assert tui._decode_key("\xe0P") == "down"
    assert tui._decode_key("\xe0K") == "left"
    assert tui._decode_key("\xe0M") == "right"


def test_navigation_state_changes(monkeypatch):
    state, calls = _make_app_state()

    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    # Navigate to the plan item (7 items down: detect, auto, targeted, cleanup, diagnostics, odt_install, odt_repair, plan)
    for _ in range(7):
        interface._handle_key("down")
    assert interface.active_tab == "plan"
    interface._handle_key("enter")
    assert interface.focus_area == "content"

    plan_cursor_before = interface.panes["plan"].cursor
    interface._handle_key("space")
    assert interface.plan_overrides["include_visio"] is True
    interface._handle_key("down")
    assert interface.panes["plan"].cursor == plan_cursor_before + 1

    interface.focus_area = "nav"
    for _ in range(len(interface.navigation)):
        if interface.navigation[interface.nav_index].name == "detect":
            break
        interface._handle_key("up")
    interface._handle_key("enter")
    assert calls["detect"] == 1


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


def test_fallback_to_cli(monkeypatch):
    state, _ = _make_app_state()

    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: False)
    called = {}

    def fake_run_cli(app_state):
        called["cli"] = True

    monkeypatch.setattr("src.office_janitor.ui.run_cli", fake_run_cli)

    interface = tui.OfficeJanitorTUI(state)
    interface.run()

    assert called.get("cli") is True


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
    """Test navigation includes ODT Install and Repair items."""
    state, _ = _make_app_state()
    monkeypatch.setattr(tui, "_supports_ansi", lambda stream=None: True)
    interface = tui.OfficeJanitorTUI(state)

    nav_names = [item.name for item in interface.navigation]
    assert "odt_install" in nav_names
    assert "odt_repair" in nav_names


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

