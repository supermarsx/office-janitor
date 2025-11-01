"""!
@brief Tests for the text-based user interface engine.
"""
from __future__ import annotations

from collections import deque
from types import SimpleNamespace

from src.office_janitor import tui


def _make_app_state():
    calls: dict[str, object] = {"detect": 0, "run": 0, "planner_overrides": None, "executor_overrides": None}

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

    for _ in range(5):
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
