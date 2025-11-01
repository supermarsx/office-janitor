"""!
@brief Tests for the text-based user interface engine.
"""
from __future__ import annotations

from collections import deque
from types import SimpleNamespace

from src.office_janitor import tui


def _make_app_state():
    calls: dict[str, int] = {"detect": 0, "run": 0}

    def detector() -> dict[str, list[str]]:
        calls["detect"] += 1
        return {"msi": ["office"]}

    def planner(inventory, overrides):
        return [{"step": "noop"}]

    def executor(plan, overrides):
        calls["run"] += 1

    state = {
        "detector": detector,
        "planner": planner,
        "executor": executor,
        "event_queue": deque(),
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
