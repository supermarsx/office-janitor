"""!
@brief Integration tests for CLI and UI layers.
@details Exercises the CLI and TUI entry points to confirm argument handling,
event emission, and execution wiring across the high-level flows.
"""

from __future__ import annotations

import argparse
import json
import pathlib
import sys

import pytest

PROJECT_ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import main, ui, version  # noqa: E402
from office_janitor import tui as tui_module  # noqa: E402
from office_janitor import tui_actions as tui_actions_module  # noqa: E402


def _no_op(*args, **kwargs):  # type: ignore[no-untyped-def]
    return None


def test_main_auto_all_executes_scrub_pipeline(monkeypatch, tmp_path) -> None:
    """!
    @brief ``--auto-all`` should run detection, planning, safety, and execution.
    """

    monkeypatch.setattr(main, "ensure_admin_and_relaunch_if_needed", _no_op)
    monkeypatch.setattr(main, "enable_vt_mode_if_possible", _no_op)
    monkeypatch.setattr(main, "_resolve_log_directory", lambda candidate: tmp_path)

    inventory = {"msi": [], "c2r": [], "filesystem": []}
    monkeypatch.setattr(main.detect, "gather_office_inventory", lambda **kw: inventory)

    recorded: list[tuple[str, object]] = []

    def fake_plan(inv, options):  # type: ignore[no-untyped-def]
        recorded.append(("plan", options))
        return [
            {
                "id": "context",
                "category": "context",
                "metadata": {
                    "mode": options.get("mode"),
                    "dry_run": options.get("dry_run"),
                    "target_versions": [],
                    "unsupported_targets": [],
                    "options": dict(options),
                },
            },
            {"id": "step-1", "category": "filesystem-cleanup", "metadata": {"paths": []}},
        ]

    monkeypatch.setattr(main.plan_module, "build_plan", fake_plan)
    monkeypatch.setattr(
        main.safety, "perform_preflight_checks", lambda plan: recorded.append(("safety", len(plan)))
    )

    scrub_calls: list[bool] = []
    monkeypatch.setattr(
        main.scrub,
        "execute_plan",
        lambda plan, dry_run=False, **kw: scrub_calls.append(bool(dry_run)),
    )

    guard_calls: list[tuple[dict, bool]] = []

    def capture_guard(options, *, dry_run=False):  # type: ignore[no-untyped-def]
        guard_calls.append((dict(options), bool(dry_run)))

    monkeypatch.setattr(main, "_enforce_runtime_guards", capture_guard)

    exit_code = main.main(["--auto-all", "--dry-run", "--logdir", str(tmp_path / "logs")])

    assert exit_code == 0
    assert scrub_calls == [True]
    assert recorded[0][0] == "plan"
    assert recorded[0][1]["mode"] == "auto-all"
    assert any(item[0] == "safety" for item in recorded)
    assert guard_calls and guard_calls[0][1] is True


def test_main_requires_confirmation_before_execution(monkeypatch, tmp_path) -> None:
    """!
    @brief Scrub execution should stop when the confirmation prompt is declined.
    """

    monkeypatch.setattr(main, "ensure_admin_and_relaunch_if_needed", _no_op)
    monkeypatch.setattr(main, "enable_vt_mode_if_possible", _no_op)
    monkeypatch.setattr(main, "_resolve_log_directory", lambda candidate: tmp_path)

    monkeypatch.setattr(main.detect, "gather_office_inventory", lambda **kw: {"msi": []})

    def fake_plan(inv, options):  # type: ignore[no-untyped-def]
        return [
            {"id": "context", "category": "context", "metadata": {"mode": options.get("mode")}},
            {"id": "step", "category": "noop", "metadata": {}},
        ]

    monkeypatch.setattr(main.plan_module, "build_plan", fake_plan)
    monkeypatch.setattr(main.safety, "perform_preflight_checks", lambda plan: None)

    guard_calls: list[tuple[dict, bool]] = []
    monkeypatch.setattr(
        main,
        "_enforce_runtime_guards",
        lambda options, *, dry_run=False: guard_calls.append((dict(options), bool(dry_run))),
    )

    scrub_calls: list[bool] = []
    monkeypatch.setattr(
        main.scrub, "execute_plan", lambda plan, dry_run=False, **kw: scrub_calls.append(True)
    )

    confirm_calls: list[dict] = []

    def deny_confirmation(**kwargs):  # type: ignore[no-untyped-def]
        confirm_calls.append(dict(kwargs))
        return False

    monkeypatch.setattr(main.confirm, "request_scrub_confirmation", deny_confirmation)

    exit_code = main.main(["--auto-all", "--logdir", str(tmp_path / "logs")])

    assert exit_code == 0
    assert confirm_calls
    assert guard_calls == []
    assert scrub_calls == []


def test_limited_user_flag_passes_to_detection(monkeypatch, tmp_path) -> None:
    """!
    @brief Limited-user flag should request de-elevated detection.
    """

    monkeypatch.setattr(main, "ensure_admin_and_relaunch_if_needed", _no_op)
    monkeypatch.setattr(main, "enable_vt_mode_if_possible", _no_op)
    monkeypatch.setattr(main, "_resolve_log_directory", lambda candidate: tmp_path)

    captured = {}

    def fake_gather(*, limited_user=None, progress_callback=None):
        captured["limited_user"] = limited_user
        return {"msi": [], "c2r": [], "filesystem": [], "registry": []}

    monkeypatch.setattr(main.detect, "gather_office_inventory", fake_gather)
    monkeypatch.setattr(main.plan_module, "build_plan", lambda inv, opts: [])
    monkeypatch.setattr(main.safety, "perform_preflight_checks", lambda plan: None)
    monkeypatch.setattr(main.confirm, "request_scrub_confirmation", lambda **kwargs: False)

    exit_code = main.main(
        ["--auto-all", "--limited-user", "--dry-run", "--logdir", str(tmp_path / "logs")]
    )

    assert exit_code == 0
    assert captured.get("limited_user") is True


def test_main_diagnose_skips_execution(monkeypatch, tmp_path) -> None:
    """!
    @brief Diagnostics mode must avoid executing the scrubber.
    """

    monkeypatch.setattr(main, "ensure_admin_and_relaunch_if_needed", _no_op)
    monkeypatch.setattr(main, "enable_vt_mode_if_possible", _no_op)
    monkeypatch.setattr(main, "_resolve_log_directory", lambda candidate: tmp_path)

    monkeypatch.setattr(
        main.detect,
        "gather_office_inventory",
        lambda **kw: {"msi": [], "c2r": [], "filesystem": []},
    )

    def fake_plan(inv, options):  # type: ignore[no-untyped-def]
        return [
            {
                "id": "context",
                "category": "context",
                "metadata": {
                    "mode": options.get("mode"),
                    "dry_run": options.get("dry_run"),
                    "target_versions": [],
                    "unsupported_targets": [],
                    "options": dict(options),
                },
            }
        ]

    monkeypatch.setattr(main.plan_module, "build_plan", fake_plan)
    monkeypatch.setattr(main.safety, "perform_preflight_checks", lambda plan: None)

    scrub_calls: list[bool] = []
    monkeypatch.setattr(
        main.scrub, "execute_plan", lambda plan, dry_run=False, **kw: scrub_calls.append(True)
    )

    guard_calls: list[tuple[dict, bool]] = []

    def capture_guard(options, *, dry_run=False):  # type: ignore[no-untyped-def]
        guard_calls.append((dict(options), bool(dry_run)))

    monkeypatch.setattr(main, "_enforce_runtime_guards", capture_guard)

    exit_code = main.main(["--diagnose", "--logdir", str(tmp_path / "logs")])

    assert exit_code == 0
    assert scrub_calls == []
    assert guard_calls == []


def test_main_interactive_uses_cli(monkeypatch, tmp_path) -> None:
    """!
    @brief Without mode flags, the plain menu should launch.
    """

    monkeypatch.setattr(main, "ensure_admin_and_relaunch_if_needed", _no_op)
    monkeypatch.setattr(main, "enable_vt_mode_if_possible", _no_op)
    monkeypatch.setattr(main, "_resolve_log_directory", lambda candidate: tmp_path)
    monkeypatch.setattr(main, "_should_use_tui", lambda args: False)

    monkeypatch.setattr(
        main,
        "_enforce_runtime_guards",
        lambda options, *, dry_run=False: (_ for _ in ()).throw(
            AssertionError("Guard should not run")
        ),
    )

    captured = {}

    def fake_run_cli(app_state):  # type: ignore[no-untyped-def]
        captured["app_state"] = app_state

    monkeypatch.setattr(main.ui, "run_cli", fake_run_cli)
    monkeypatch.setattr(
        main.tui,
        "run_tui",
        lambda app_state: (_ for _ in ()).throw(AssertionError("TUI not expected")),
    )

    exit_code = main.main(["--logdir", str(tmp_path / "logs")])

    assert exit_code == 0
    assert "detector" in captured["app_state"]


def test_ui_plan_and_execute_skips_without_confirmation(monkeypatch) -> None:
    """!
    @brief Interactive runs should abort when confirmation is denied.
    """

    monkeypatch.setattr(ui.plan_module, "summarize_plan", lambda plan: {"total_steps": len(plan)})

    executed: list[tuple[list, dict]] = []

    context = {
        "detector": lambda: {},
        "planner": lambda inventory, overrides: [
            {"id": "step", "category": "noop", "metadata": {}}
        ],
        "executor": lambda plan, overrides: executed.append((plan, dict(overrides or {}))),
        "args": argparse.Namespace(dry_run=False, force=False),
        "human_logger": None,
        "machine_logger": None,
        "emit_event": None,
        "event_queue": None,
        "input": lambda prompt: "n",
        "confirm": lambda **kwargs: False,
        "inventory": {},
        "plan": None,
        "running": True,
    }

    ui._plan_and_execute(context, {"mode": "auto-all", "auto_all": True}, label="auto scrub")

    assert executed == []


def test_ui_plan_and_execute_runs_after_confirmation(monkeypatch) -> None:
    """!
    @brief Interactive runs should proceed when confirmation is accepted.
    """

    monkeypatch.setattr(ui.plan_module, "summarize_plan", lambda plan: {"total_steps": len(plan)})

    executed: list[tuple[list, dict]] = []

    def record_executor(plan, overrides):  # type: ignore[no-untyped-def]
        executed.append((plan, dict(overrides or {})))

    def fake_input(prompt: str) -> str:
        return "Y"

    context = {
        "detector": lambda: {},
        "planner": lambda inventory, overrides: [
            {"id": "step", "category": "noop", "metadata": {}}
        ],
        "executor": record_executor,
        "args": argparse.Namespace(dry_run=False, force=False),
        "human_logger": None,
        "machine_logger": None,
        "emit_event": None,
        "event_queue": None,
        "input": fake_input,
        "confirm": lambda **kwargs: True,
        "inventory": {},
        "plan": None,
        "running": True,
    }

    ui._plan_and_execute(context, {"mode": "auto-all", "auto_all": True}, label="auto scrub")

    assert executed
    plan_override = executed[0][1]
    assert plan_override.get("confirmed") is True
    assert plan_override.get("interactive") is True
    assert callable(plan_override.get("input_func"))


def test_arg_parser_and_plan_options_cover_modes() -> None:
    """!
    @brief ``build_arg_parser`` should expose every documented switch.
    """

    parser = main.build_arg_parser()
    args = parser.parse_args(
        [
            "--target",
            "2016",
            "--include",
            "visio,project",
            "--force",
            "--allow-unsupported-windows",
            "--dry-run",
            "--no-restore-point",
            "--no-license",
            "--keep-templates",
            "--timeout",
            "90",
            "--backup",
            "C:/backup",
        ]
    )

    mode = main._determine_mode(args)
    options = main._collect_plan_options(args, mode)

    assert mode == "target:2016"
    assert options["target"] == "2016"
    assert options["include"] == "visio,project"
    assert options["force"] is True
    assert options["allow_unsupported_windows"] is True
    assert options["dry_run"] is True
    assert options["create_restore_point"] is False
    assert options["no_license"] is True
    assert options["keep_templates"] is True
    assert options["timeout"] == 90
    assert options["backup"] == "C:/backup"


def test_version_option_reports_metadata(capsys) -> None:
    """!
    @brief ``--version`` should surface both version and build identifiers.
    """

    parser = main.build_arg_parser()
    with pytest.raises(SystemExit):
        parser.parse_args(["--version"])
    output = capsys.readouterr().out
    info = version.build_info()
    assert info["version"] in output
    assert info["build"] in output


def test_ui_header_displays_build_info(capsys) -> None:
    """!
    @brief The interactive menu header should mention version metadata.
    """

    ui._print_menu([])
    output = capsys.readouterr().out
    info = version.build_info()
    assert info["version"] in output
    assert info["build"] in output


def test_main_target_mode_passes_all_options(monkeypatch, tmp_path) -> None:
    """!
    @brief ``--target`` should propagate all ancillary options into the plan.
    """

    monkeypatch.setattr(main, "ensure_admin_and_relaunch_if_needed", _no_op)
    monkeypatch.setattr(main, "enable_vt_mode_if_possible", _no_op)
    monkeypatch.setattr(main, "_resolve_log_directory", lambda candidate: tmp_path)

    inventory = {"msi": ["Office16"], "c2r": [], "filesystem": []}
    monkeypatch.setattr(main.detect, "gather_office_inventory", lambda **kw: inventory)

    captured_options: dict = {}

    def fake_plan(inv, options):  # type: ignore[no-untyped-def]
        captured_options.update(options)
        return [
            {
                "id": "context",
                "category": "context",
                "metadata": {"mode": options.get("mode"), "options": dict(options)},
            },
            {
                "id": "filesystem",
                "category": "filesystem-cleanup",
                "metadata": {"paths": ["C:/Office"]},
            },
        ]

    monkeypatch.setattr(main.plan_module, "build_plan", fake_plan)
    monkeypatch.setattr(main.safety, "perform_preflight_checks", lambda plan: None)

    scrub_calls: list[bool] = []

    def fake_execute(plan, dry_run=False, **kw):  # type: ignore[no-untyped-def]
        scrub_calls.append(bool(dry_run))

    monkeypatch.setattr(main.scrub, "execute_plan", fake_execute)

    guard_calls: list[tuple[dict, bool]] = []

    def capture_guard(options, *, dry_run=False):  # type: ignore[no-untyped-def]
        guard_calls.append((dict(options), bool(dry_run)))

    monkeypatch.setattr(main, "_enforce_runtime_guards", capture_guard)

    plan_path = tmp_path / "plan.json"
    backup_dir = tmp_path / "backup"
    exit_code = main.main(
        [
            "--target",
            "2016",
            "--include",
            "visio,project",
            "--force",
            "--allow-unsupported-windows",
            "--keep-templates",
            "--no-license",
            "--timeout",
            "120",
            "--backup",
            str(backup_dir),
            "--plan",
            str(plan_path),
            "--no-restore-point",
            "--logdir",
            str(tmp_path / "logs"),
        ]
    )

    assert exit_code == 0
    assert scrub_calls == [False]
    assert captured_options["mode"] == "target:2016"
    assert captured_options["target"] == "2016"
    assert captured_options["include"] == "visio,project"
    assert captured_options["force"] is True
    assert captured_options["allow_unsupported_windows"] is True
    assert captured_options["keep_templates"] is True
    assert captured_options["no_license"] is True
    assert captured_options["timeout"] == 120
    assert captured_options["create_restore_point"] is False
    assert (backup_dir / "plan.json").exists()
    assert (backup_dir / "inventory.json").exists()
    assert plan_path.exists()
    assert guard_calls
    guard_options, guard_dry_run = guard_calls[0]
    assert guard_options["allow_unsupported_windows"] is True
    assert guard_options["force"] is True
    assert guard_dry_run is False


def test_ui_run_cli_detect_option(monkeypatch) -> None:
    """!
    @brief Menu option 1 should call the detector and exit cleanly.
    """

    events: list[str] = []
    inputs = iter(["1", "10"])

    def fake_input(prompt: str) -> str:
        return next(inputs)

    app_state = {
        "args": type(
            "Args",
            (),
            {
                "quiet": False,
                "dry_run": False,
                "no_restore_point": False,
                "logdir": "logs",
                "backup": None,
            },
        )(),
        "detector": lambda: events.append("detect") or {"msi": [1], "c2r": [], "filesystem": []},
        "planner": lambda inventory, overrides=None: (_ for _ in ()).throw(
            AssertionError("planner not expected")
        ),
        "executor": lambda plan, overrides=None: (_ for _ in ()).throw(
            AssertionError("executor not expected")
        ),
        "input": fake_input,
        "confirm": lambda **kwargs: True,
    }

    ui.run_cli(app_state)

    assert events == ["detect"]


def test_ui_run_cli_auto_all_executes(monkeypatch) -> None:
    """!
    @brief Menu option 2 should plan and execute using overrides.
    """

    events: list[tuple[str, object]] = []
    inputs = iter(["2", "10"])

    def fake_input(prompt: str) -> str:
        return next(inputs)

    def fake_detector():
        events.append(("detect", None))
        return {"msi": [], "c2r": [], "filesystem": []}

    def fake_planner(inventory, overrides=None):
        events.append(("plan", overrides))
        return [
            {
                "id": "context",
                "category": "context",
                "metadata": {
                    "mode": overrides.get("mode") if overrides else "interactive",
                    "dry_run": False,
                    "target_versions": [],
                    "unsupported_targets": [],
                    "options": {},
                },
            },
            {"id": "registry-0", "category": "registry-cleanup", "metadata": {"keys": []}},
        ]

    def fake_executor(plan, overrides=None):
        events.append(("execute", overrides))

    app_state = {
        "args": type(
            "Args",
            (),
            {
                "quiet": False,
                "dry_run": False,
                "no_restore_point": False,
                "logdir": "logs",
                "backup": None,
            },
        )(),
        "detector": fake_detector,
        "planner": fake_planner,
        "executor": fake_executor,
        "input": fake_input,
        "confirm": lambda **kwargs: True,
    }

    ui.run_cli(app_state)

    assert events[0][0] == "detect"
    assert events[1][0] == "plan"
    assert events[1][1]["mode"] == "auto-all"
    assert events[1][1]["auto_all"] is True
    assert events[2][0] == "execute"
    assert events[2][1]["mode"] == "auto-all"
    assert events[2][1]["auto_all"] is True
    assert "inventory" in events[2][1]


def test_ui_run_cli_targeted_prompts(monkeypatch) -> None:
    """!
    @brief Menu option 3 should collect target versions and includes.
    """

    events: list[tuple[str, object]] = []
    inputs = iter(["3", "2016,365", "visio,project", "10"])

    def fake_input(prompt: str) -> str:
        return next(inputs)

    def fake_detector():
        events.append(("detect", None))
        return {"msi": [], "c2r": [], "filesystem": []}

    def fake_planner(inventory, overrides=None):
        events.append(("plan", overrides))
        return [
            {
                "id": "context",
                "category": "context",
                "metadata": {
                    "mode": overrides.get("mode") if overrides else "interactive",
                    "dry_run": False,
                    "target_versions": [],
                    "unsupported_targets": [],
                    "options": {},
                },
            }
        ]

    def fake_executor(plan, overrides=None):
        events.append(("execute", overrides))

    app_state = {
        "args": type(
            "Args",
            (),
            {
                "quiet": False,
                "dry_run": False,
                "no_restore_point": False,
                "logdir": "logs",
                "backup": None,
            },
        )(),
        "detector": fake_detector,
        "planner": fake_planner,
        "executor": fake_executor,
        "input": fake_input,
        "confirm": lambda **kwargs: True,
    }

    ui.run_cli(app_state)

    assert events[0][0] == "detect"
    assert events[1][1]["mode"] == "target:2016,365"
    assert events[1][1]["target"] == "2016,365"
    assert events[1][1]["include"] == "visio,project"
    assert events[2][1]["mode"] == "target:2016,365"


def test_ui_run_cli_respects_json_flag() -> None:
    """!
    @brief Menu should not launch when ``--json`` is requested.
    """

    events: list[str] = []

    app_state = {
        "args": type(
            "Args",
            (),
            {
                "quiet": False,
                "json": True,
                "dry_run": False,
                "no_restore_point": False,
                "logdir": "logs",
                "backup": None,
            },
        )(),
        "detector": lambda: events.append("detect"),
        "planner": lambda inventory, overrides=None: events.append("plan"),
        "executor": lambda plan, overrides=None: events.append("execute"),
        "confirm": lambda **kwargs: True,
    }

    ui.run_cli(app_state)

    assert events == []


def test_tui_falls_back_without_ansi(monkeypatch, capsys) -> None:
    """!
    @brief When ANSI support is missing the TUI should print an error and return.
    """

    monkeypatch.setattr(tui_module, "_supports_ansi", lambda stream=None: False)

    tui_module.run_tui(
        {
            "args": type("Args", (), {"no_color": False, "quiet": False})(),
            "detector": lambda: {},
            "planner": lambda inventory, overrides=None: [],
            "executor": lambda plan, overrides=None: None,
            "confirm": lambda **kwargs: True,
        }
    )

    captured = capsys.readouterr()
    assert "ANSI terminal support" in captured.out


def test_tui_commands_drive_backends(monkeypatch) -> None:
    """!
    @brief Key commands should call detector, planner, and executor in order.
    """

    monkeypatch.setattr(tui_module, "_supports_ansi", lambda stream=None: True)
    monkeypatch.setattr(tui_module, "_spinner", lambda duration, message: None)
    monkeypatch.setattr(tui_actions_module, "spinner", lambda duration, message: None)
    monkeypatch.setattr(tui_module.OfficeJanitorTUI, "_drain_events", lambda self: False)

    keys = iter(
        [
            # Mode selection: select "remove" (3rd option, index 2)
            "down",  # repair
            "down",  # remove
            "enter",  # select remove mode
            # Now in remove mode, go to auto and trigger it
            "down",  # auto remove all
            "enter",  # trigger auto (which does detect, plan, execute)
            "quit",
        ]
    )

    def reader() -> str:
        return next(keys)

    events: list[str] = []

    def fake_detector():
        events.append("detect")
        return {"msi": [], "c2r": [], "filesystem": []}

    def fake_planner(inventory, overrides=None):
        events.append("plan")
        return [
            {
                "id": "context",
                "category": "context",
                "metadata": {
                    "mode": "interactive",
                    "dry_run": False,
                    "target_versions": [],
                    "unsupported_targets": [],
                    "options": {},
                },
            },
            {"id": "filesystem-0", "category": "filesystem-cleanup", "metadata": {"paths": []}},
        ]

    def fake_executor(plan, overrides=None):
        events.append("execute")

    tui_module.run_tui(
        {
            "args": type("Args", (), {"no_color": False, "quiet": False})(),
            "detector": fake_detector,
            "planner": fake_planner,
            "executor": fake_executor,
            "key_reader": reader,
            "confirm": lambda **kwargs: True,
        }
    )

    assert events == ["detect", "plan", "execute"]


def test_tui_respects_quiet_flag(monkeypatch) -> None:
    """!
    @brief Quiet mode should prevent the TUI from starting an event loop.
    """

    monkeypatch.setattr(tui_module, "_supports_ansi", lambda stream=None: True)

    invoked: list[str] = []

    def reader() -> str:
        invoked.append("reader")
        return "q"

    tui_module.run_tui(
        {
            "args": type("Args", (), {"no_color": False, "quiet": True, "json": False})(),
            "detector": lambda: invoked.append("detect"),
            "planner": lambda inventory, overrides=None: invoked.append("plan") or [],
            "executor": lambda plan, overrides=None: invoked.append("execute"),
            "key_reader": reader,
            "confirm": lambda **kwargs: True,
        }
    )

    assert invoked == []


def test_tui_auto_mode_invokes_overrides(monkeypatch) -> None:
    """!
    @brief The ``A`` command should trigger auto scrub with overrides.
    """

    monkeypatch.setattr(tui_module, "_supports_ansi", lambda stream=None: True)
    monkeypatch.setattr(tui_module, "_spinner", lambda duration, message: None)
    monkeypatch.setattr(tui_actions_module, "spinner", lambda duration, message: None)
    monkeypatch.setattr(tui_module.OfficeJanitorTUI, "_drain_events", lambda self: False)

    keys = iter([
        # Mode selection: select "remove" (3rd option, index 2)
        "down",  # repair
        "down",  # remove
        "enter",  # select remove mode
        # Now in remove mode, navigate to auto
        "down",  # auto remove all (2nd item)
        "enter",  # trigger auto
        "quit",
    ])

    def reader() -> str:
        return next(keys)

    events: list[tuple[str, object]] = []

    def fake_detector():
        events.append(("detect", None))
        return {"msi": [], "c2r": [], "filesystem": []}

    def fake_planner(inventory, overrides=None):
        events.append(("plan", overrides))
        return [
            {
                "id": "context",
                "category": "context",
                "metadata": {"mode": overrides.get("mode"), "options": dict(overrides)},
            }
        ]

    def fake_executor(plan, overrides=None):
        events.append(("execute", overrides))

    tui_module.run_tui(
        {
            "args": type("Args", (), {"no_color": False, "quiet": False})(),
            "detector": fake_detector,
            "planner": fake_planner,
            "executor": fake_executor,
            "key_reader": reader,
            "confirm": lambda **kwargs: True,
        }
    )

    assert events[0][0] == "detect"
    assert events[1][1]["mode"] == "auto-all"
    assert events[1][1]["auto_all"] is True
    assert events[2][1]["auto_all"] is True


def test_tui_targeted_collects_input(monkeypatch) -> None:
    """!
    @brief The ``T`` command should prompt for versions and includes.
    """

    monkeypatch.setattr(tui_module, "_supports_ansi", lambda stream=None, **_: True)
    monkeypatch.setattr(tui_module, "_spinner", lambda duration, message: None)
    monkeypatch.setattr(tui_actions_module, "spinner", lambda duration, message: None)
    monkeypatch.setattr(tui_module.OfficeJanitorTUI, "_drain_events", lambda self: False)
    monkeypatch.setattr(
        tui_module.OfficeJanitorTUI, "_selected_targets", lambda self: ["2016", "365"]
    )
    monkeypatch.setattr(
        tui_module.OfficeJanitorTUI, "_collect_plan_overrides", lambda self: {"include": "visio"}
    )

    keys = iter([
        # Mode selection: select "remove" (3rd option, index 2)
        "down",  # repair
        "down",  # remove
        "enter",  # select remove mode
        # Navigate to targeted remove (3rd item)
        "down",  # auto
        "down",  # targeted
        "enter",  # activate targeted
        "f10",   # run targeted
        "quit",
    ])

    def reader() -> str:
        return next(keys)

    prompts = iter(["2016,365", "visio"])

    monkeypatch.setattr(tui_module, "_read_input_line", lambda prompt: next(prompts))

    events: list[tuple[str, object]] = []

    def fake_detector():
        events.append(("detect", None))
        return {"msi": [], "c2r": [], "filesystem": []}

    def fake_planner(inventory, overrides=None):
        events.append(("plan", overrides))
        return [
            {
                "id": "context",
                "category": "context",
                "metadata": {"mode": overrides.get("mode"), "options": dict(overrides)},
            }
        ]

    def fake_executor(plan, overrides=None):
        events.append(("execute", overrides))

    tui_module.run_tui(
        {
            "args": type("Args", (), {"no_color": False, "quiet": False})(),
            "detector": fake_detector,
            "planner": fake_planner,
            "executor": fake_executor,
            "key_reader": reader,
            "confirm": lambda **kwargs: True,
        }
    )

    assert events[0][0] == "detect"
    assert events[1][1]["mode"] == "target:2016,365"
    assert events[1][1]["include"] == "visio"
    assert events[2][1]["target"] == "2016,365"


def test_main_diagnose_writes_default_artifacts(monkeypatch, tmp_path) -> None:
    """!
    @brief Diagnostics mode should persist plan and inventory to the log directory.
    """

    monkeypatch.setattr(main, "ensure_admin_and_relaunch_if_needed", _no_op)
    monkeypatch.setattr(main, "enable_vt_mode_if_possible", _no_op)
    monkeypatch.setattr(main, "_resolve_log_directory", lambda candidate: tmp_path)

    monkeypatch.setattr(
        main.detect,
        "gather_office_inventory",
        lambda **kw: {"msi": ["Office"], "c2r": [], "filesystem": []},
    )

    def fake_plan(inv, options):  # type: ignore[no-untyped-def]
        return [
            {
                "id": "context",
                "category": "context",
                "metadata": {"mode": options.get("mode"), "options": dict(options)},
            },
            {
                "id": "filesystem",
                "category": "filesystem-cleanup",
                "metadata": {"paths": ["C:/Office"]},
            },
        ]

    monkeypatch.setattr(main.plan_module, "build_plan", fake_plan)
    monkeypatch.setattr(main.safety, "perform_preflight_checks", lambda plan: None)
    monkeypatch.setattr(main.scrub, "execute_plan", lambda plan, dry_run=False, **kw: None)

    guard_calls: list[tuple[dict, bool]] = []

    def capture_guard(options, *, dry_run=False):  # type: ignore[no-untyped-def]
        guard_calls.append((dict(options), bool(dry_run)))

    monkeypatch.setattr(main, "_enforce_runtime_guards", capture_guard)

    exit_code = main.main(["--diagnose", "--logdir", str(tmp_path)])

    assert exit_code == 0
    plan_path = tmp_path / "diagnostics-plan.json"
    inventory_path = tmp_path / "diagnostics-inventory.json"
    assert plan_path.exists()
    assert inventory_path.exists()
    data = json.loads(plan_path.read_text(encoding="utf-8"))
    assert data[0]["metadata"]["mode"] == "diagnose"
    inventory_data = json.loads(inventory_path.read_text(encoding="utf-8"))
    assert inventory_data["msi"] == ["Office"]
    assert guard_calls == []


# ===========================================================================
# Comprehensive CLI argument tests
# ===========================================================================


class TestCLIArgumentParsing:
    """!
    @brief Test all CLI arguments are correctly parsed and handled.
    """

    def test_version_short_flag(self, capsys) -> None:
        """Test -V flag outputs version info."""
        parser = main.build_arg_parser()
        with pytest.raises(SystemExit) as exc_info:
            parser.parse_args(["-V"])
        assert exc_info.value.code == 0
        output = capsys.readouterr().out
        assert "0." in output or "dev" in output  # Version format

    def test_version_long_flag(self, capsys) -> None:
        """Test --version flag outputs version info."""
        parser = main.build_arg_parser()
        with pytest.raises(SystemExit) as exc_info:
            parser.parse_args(["--version"])
        assert exc_info.value.code == 0

    def test_help_flag(self, capsys) -> None:
        """Test --help flag outputs usage info."""
        parser = main.build_arg_parser()
        with pytest.raises(SystemExit) as exc_info:
            parser.parse_args(["--help"])
        assert exc_info.value.code == 0
        output = capsys.readouterr().out
        assert "office-janitor" in output
        assert "--auto-all" in output
        assert "--diagnose" in output

    def test_auto_all_mode(self) -> None:
        """Test --auto-all sets the mode correctly."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all"])
        assert args.auto_all is True
        assert main._determine_mode(args) == "auto-all"

    def test_target_mode_with_version(self) -> None:
        """Test --target sets the mode and version correctly."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--target", "2019"])
        assert args.target == "2019"
        assert main._determine_mode(args) == "target:2019"

    def test_target_mode_with_multiple_versions(self) -> None:
        """Test --target accepts comma-separated versions."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--target", "2016,2019,365"])
        assert args.target == "2016,2019,365"
        assert main._determine_mode(args) == "target:2016,2019,365"

    def test_diagnose_mode(self) -> None:
        """Test --diagnose sets the mode correctly."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--diagnose"])
        assert args.diagnose is True
        assert main._determine_mode(args) == "diagnose"

    def test_cleanup_only_mode(self) -> None:
        """Test --cleanup-only sets the mode correctly."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--cleanup-only"])
        assert args.cleanup_only is True
        assert main._determine_mode(args) == "cleanup-only"

    def test_interactive_mode_default(self) -> None:
        """Test interactive mode is default when no mode flags provided."""
        parser = main.build_arg_parser()
        args = parser.parse_args([])
        assert main._determine_mode(args) == "interactive"

    def test_modes_are_mutually_exclusive(self, capsys) -> None:
        """Test that mode flags are mutually exclusive."""
        parser = main.build_arg_parser()
        with pytest.raises(SystemExit) as exc_info:
            parser.parse_args(["--auto-all", "--diagnose"])
        assert exc_info.value.code != 0

    def test_include_components(self) -> None:
        """Test --include accepts component list."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--include", "visio,project,access"])
        assert args.include == "visio,project,access"

    def test_force_flag(self) -> None:
        """Test --force flag."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--force"])
        assert args.force is True

    def test_allow_unsupported_windows_flag(self) -> None:
        """Test --allow-unsupported-windows flag."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--allow-unsupported-windows"])
        assert args.allow_unsupported_windows is True

    def test_dry_run_flag(self) -> None:
        """Test --dry-run flag."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--dry-run"])
        assert args.dry_run is True

    def test_no_restore_point_flag(self) -> None:
        """Test --no-restore-point flag."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--no-restore-point"])
        assert args.no_restore_point is True

    def test_no_license_flag(self) -> None:
        """Test --no-license flag."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--no-license"])
        assert args.no_license is True

    def test_keep_license_flag(self) -> None:
        """Test --keep-license flag (alias of --no-license)."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--keep-license"])
        assert args.keep_license is True

    def test_keep_templates_flag(self) -> None:
        """Test --keep-templates flag."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--keep-templates"])
        assert args.keep_templates is True

    def test_plan_output_path(self) -> None:
        """Test --plan accepts output path."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--plan", "C:/output/plan.json"])
        assert args.plan == "C:/output/plan.json"

    def test_logdir_path(self) -> None:
        """Test --logdir accepts directory path."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--logdir", "C:/logs"])
        assert args.logdir == "C:/logs"

    def test_backup_path(self) -> None:
        """Test --backup accepts directory path."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--backup", "C:/backup"])
        assert args.backup == "C:/backup"

    def test_timeout_seconds(self) -> None:
        """Test --timeout accepts integer seconds."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--timeout", "120"])
        assert args.timeout == 120

    def test_timeout_requires_integer(self, capsys) -> None:
        """Test --timeout rejects non-integer values."""
        parser = main.build_arg_parser()
        with pytest.raises(SystemExit) as exc_info:
            parser.parse_args(["--timeout", "not-a-number"])
        assert exc_info.value.code != 0

    def test_quiet_flag(self) -> None:
        """Test --quiet flag."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--quiet"])
        assert args.quiet is True

    def test_json_flag(self) -> None:
        """Test --json flag."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--json"])
        assert args.json is True

    def test_tui_flag(self) -> None:
        """Test --tui flag."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--tui"])
        assert args.tui is True

    def test_no_color_flag(self) -> None:
        """Test --no-color flag."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--no-color"])
        assert args.no_color is True

    def test_tui_compact_flag(self) -> None:
        """Test --tui-compact flag."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--tui-compact"])
        assert args.tui_compact is True

    def test_tui_refresh_milliseconds(self) -> None:
        """Test --tui-refresh accepts milliseconds."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--tui-refresh", "500"])
        assert args.tui_refresh == 500

    def test_limited_user_flag(self) -> None:
        """Test --limited-user flag."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--limited-user"])
        assert args.limited_user is True

    def test_all_flags_combined(self) -> None:
        """Test multiple flags can be combined."""
        parser = main.build_arg_parser()
        args = parser.parse_args(
            [
                "--auto-all",
                "--force",
                "--dry-run",
                "--no-restore-point",
                "--no-license",
                "--keep-templates",
                "--quiet",
                "--json",
                "--no-color",
                "--allow-unsupported-windows",
                "--limited-user",
                "--include",
                "visio",
                "--timeout",
                "60",
                "--logdir",
                "/tmp/logs",
                "--backup",
                "/tmp/backup",
                "--plan",
                "/tmp/plan.json",
            ]
        )
        assert args.auto_all is True
        assert args.force is True
        assert args.dry_run is True
        assert args.no_restore_point is True
        assert args.no_license is True
        assert args.keep_templates is True
        assert args.quiet is True
        assert args.json is True
        assert args.no_color is True
        assert args.allow_unsupported_windows is True
        assert args.limited_user is True
        assert args.include == "visio"
        assert args.timeout == 60
        assert args.logdir == "/tmp/logs"
        assert args.backup == "/tmp/backup"
        assert args.plan == "/tmp/plan.json"


class TestCLIArgumentsIntoPlanOptions:
    """!
    @brief Test CLI arguments are correctly propagated to plan options.
    """

    def test_force_in_plan_options(self) -> None:
        """Test --force propagates to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--force"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["force"] is True

    def test_dry_run_in_plan_options(self) -> None:
        """Test --dry-run propagates to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--dry-run"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["dry_run"] is True

    def test_no_restore_point_in_plan_options(self) -> None:
        """Test --no-restore-point sets create_restore_point=False."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--no-restore-point"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["create_restore_point"] is False

    def test_restore_point_default(self) -> None:
        """Test create_restore_point defaults to True."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["create_restore_point"] is True

    def test_no_license_in_plan_options(self) -> None:
        """Test --no-license propagates to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--no-license"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["no_license"] is True

    def test_keep_license_in_plan_options(self) -> None:
        """Test --keep-license sets no_license=True in plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--keep-license"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["keep_license"] is True

    def test_keep_templates_in_plan_options(self) -> None:
        """Test --keep-templates propagates to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--keep-templates"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["keep_templates"] is True

    def test_target_in_plan_options(self) -> None:
        """Test --target value propagates to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--target", "2019"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["target"] == "2019"

    def test_include_in_plan_options(self) -> None:
        """Test --include value propagates to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--include", "visio,project"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["include"] == "visio,project"

    def test_timeout_in_plan_options(self) -> None:
        """Test --timeout value propagates to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--timeout", "90"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["timeout"] == 90

    def test_backup_in_plan_options(self) -> None:
        """Test --backup value propagates to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--backup", "C:/backup"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["backup"] == "C:/backup"

    def test_allow_unsupported_windows_in_plan_options(self) -> None:
        """Test --allow-unsupported-windows propagates to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--allow-unsupported-windows"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["allow_unsupported_windows"] is True

    def test_mode_in_plan_options(self) -> None:
        """Test mode is included in plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--diagnose"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["mode"] == "diagnose"

    # --- New tests for extended CLI options ---

    def test_uninstall_method_in_plan_options(self) -> None:
        """Test --uninstall-method propagates to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--uninstall-method", "odt"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["uninstall_method"] == "odt"

    def test_msi_only_sets_uninstall_method(self) -> None:
        """Test --msi-only sets uninstall_method to msi."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--msi-only"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["uninstall_method"] == "msi"

    def test_c2r_only_sets_uninstall_method(self) -> None:
        """Test --c2r-only sets uninstall_method to c2r."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--c2r-only"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["uninstall_method"] == "c2r"

    def test_force_app_shutdown_in_plan_options(self) -> None:
        """Test --force-app-shutdown propagates to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--force-app-shutdown"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["force_app_shutdown"] is True

    def test_product_codes_in_plan_options(self) -> None:
        """Test --product-code propagates to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(
            [
                "--auto-all",
                "--product-code",
                "{00000000-0000-0000-0000-000000000001}",
                "--product-code",
                "{00000000-0000-0000-0000-000000000002}",
            ]
        )
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["product_codes"] == [
            "{00000000-0000-0000-0000-000000000001}",
            "{00000000-0000-0000-0000-000000000002}",
        ]

    def test_release_ids_in_plan_options(self) -> None:
        """Test --release-id propagates to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(
            [
                "--auto-all",
                "--release-id",
                "O365ProPlusRetail",
                "--release-id",
                "VisioProRetail",
            ]
        )
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["release_ids"] == ["O365ProPlusRetail", "VisioProRetail"]

    def test_scrub_level_in_plan_options(self) -> None:
        """Test --scrub-level propagates to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--scrub-level", "aggressive"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["scrub_level"] == "aggressive"

    def test_max_passes_in_plan_options(self) -> None:
        """Test --max-passes propagates to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--max-passes", "5"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["max_passes"] == 5

    def test_passes_in_plan_options(self) -> None:
        """Test --passes propagates to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--passes", "3"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["max_passes"] == 3

    def test_passes_takes_precedence_over_max_passes(self) -> None:
        """Test --passes takes precedence over --max-passes."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--max-passes", "5", "--passes", "2"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["max_passes"] == 2

    def test_default_passes_is_one(self) -> None:
        """Test default passes is 1 (not 3 like before)."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["max_passes"] == 1

    def test_skip_uninstall_sets_passes_zero(self) -> None:
        """Test --skip-uninstall sets max_passes to 0."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--skip-uninstall"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["max_passes"] == 0

    def test_skip_uninstall_overrides_passes(self) -> None:
        """Test --skip-uninstall takes precedence over --passes."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--passes", "5", "--skip-uninstall"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["max_passes"] == 0

    def test_registry_only_sets_correct_options(self) -> None:
        """Test --registry-only sets passes=0, skip_filesystem, no_license, etc."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--registry-only"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["max_passes"] == 0
        assert options["skip_filesystem"] is True
        assert options["skip_processes"] is True
        assert options["skip_services"] is True
        assert options["no_license"] is True
        assert options["registry_only"] is True
        # Registry should NOT be skipped
        assert options["skip_registry"] is False

    def test_skip_flags_in_plan_options(self) -> None:
        """Test skip flags propagate to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(
            [
                "--auto-all",
                "--skip-processes",
                "--skip-services",
                "--skip-tasks",
                "--skip-registry",
                "--skip-filesystem",
            ]
        )
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["skip_processes"] is True
        assert options["skip_services"] is True
        assert options["skip_tasks"] is True
        assert options["skip_registry"] is True
        assert options["skip_filesystem"] is True

    def test_clean_flags_in_plan_options(self) -> None:
        """Test clean flags propagate to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(
            [
                "--auto-all",
                "--clean-msocache",
                "--clean-appx",
                "--clean-wi-metadata",
            ]
        )
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["clean_msocache"] is True
        assert options["clean_appx"] is True
        assert options["clean_wi_metadata"] is True

    def test_license_flags_in_plan_options(self) -> None:
        """Test license cleanup flags propagate to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(
            [
                "--auto-all",
                "--clean-spp",
                "--clean-ospp",
                "--clean-vnext",
            ]
        )
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["clean_spp"] is True
        assert options["clean_ospp"] is True
        assert options["clean_vnext"] is True

    def test_user_data_flags_in_plan_options(self) -> None:
        """Test user data flags propagate to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(
            [
                "--auto-all",
                "--keep-user-settings",
                "--keep-outlook-data",
                "--keep-outlook-signatures",
                "--clean-shortcuts",
            ]
        )
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["keep_user_settings"] is True
        assert options["keep_outlook_data"] is True
        assert options["keep_outlook_signatures"] is True
        assert options["clean_shortcuts"] is True

    def test_registry_cleanup_flags_in_plan_options(self) -> None:
        """Test registry cleanup flags propagate to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(
            [
                "--auto-all",
                "--clean-addin-registry",
                "--clean-com-registry",
                "--clean-shell-extensions",
                "--clean-typelibs",
                "--clean-protocol-handlers",
                "--remove-vba",
            ]
        )
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["clean_addin_registry"] is True
        assert options["clean_com_registry"] is True
        assert options["clean_shell_extensions"] is True
        assert options["clean_typelibs"] is True
        assert options["clean_protocol_handlers"] is True
        assert options["remove_vba"] is True

    def test_retry_options_in_plan_options(self) -> None:
        """Test retry options propagate to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(
            [
                "--auto-all",
                "--retries",
                "5",
                "--retry-delay",
                "10",
                "--retry-delay-max",
                "60",
            ]
        )
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["retries"] == 5
        assert options["retry_delay"] == 10
        assert options["retry_delay_max"] == 60

    def test_no_reboot_in_plan_options(self) -> None:
        """Test --no-reboot propagates to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--no-reboot"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["no_reboot"] is True

    def test_offline_in_plan_options(self) -> None:
        """Test --offline propagates to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--offline"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["offline"] is True

    def test_advanced_flags_in_plan_options(self) -> None:
        """Test advanced flags propagate to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(
            [
                "--auto-all",
                "--skip-preflight",
                "--skip-backup",
                "--skip-verification",
                "--schedule-reboot",
            ]
        )
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["skip_preflight"] is True
        assert options["skip_backup"] is True
        assert options["skip_verification"] is True
        assert options["schedule_reboot"] is True

    def test_msiexec_args_in_plan_options(self) -> None:
        """Test --msiexec-args propagates to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--msiexec-args", "/l*v C:/log.txt"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["msiexec_args"] == "/l*v C:/log.txt"

    def test_offscrub_flags_in_plan_options(self) -> None:
        """Test OffScrub legacy flags propagate to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(
            [
                "--auto-all",
                "--offscrub-all",
                "--offscrub-ose",
                "--offscrub-offline",
                "--offscrub-quiet",
                "--offscrub-test-rerun",
            ]
        )
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["offscrub_all"] is True
        assert options["offscrub_ose"] is True
        assert options["offscrub_offline"] is True
        assert options["offscrub_quiet"] is True
        assert options["offscrub_test_rerun"] is True

    def test_verbose_in_plan_options(self) -> None:
        """Test -v flags propagate to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "-vvv"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["verbose"] == 3

    def test_yes_flag_in_plan_options(self) -> None:
        """Test --yes flag propagates to plan options."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--yes"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        assert options["yes"] is True

    def test_auto_all_enables_comprehensive_scrubbing(self) -> None:
        """Test --auto-all enables all aggressive scrub options by default."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        # Auto-all should set nuclear scrub level
        assert options["scrub_level"] == "nuclear"
        # Force options
        assert options["force_app_shutdown"] is True
        assert options["force"] is True
        # Clean everything
        assert options["clean_msocache"] is True
        assert options["clean_appx"] is True
        assert options["clean_wi_metadata"] is True
        # Clean all licenses
        assert options["clean_spp"] is True
        assert options["clean_ospp"] is True
        assert options["clean_vnext"] is True
        assert options["clean_all_licenses"] is True
        # Clean registry
        assert options["clean_addin_registry"] is True
        assert options["clean_com_registry"] is True
        assert options["clean_shell_extensions"] is True
        assert options["clean_typelibs"] is True
        assert options["clean_protocol_handlers"] is True
        # Clean user data
        assert options["clean_shortcuts"] is True
        assert options["delete_user_settings"] is True
        # VBA
        assert options["remove_vba"] is True
        # OffScrub-style options
        assert options["offscrub_all"] is True
        assert options["offscrub_ose"] is True

    def test_auto_all_respects_explicit_scrub_level(self) -> None:
        """Test --auto-all respects user-specified --scrub-level."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--scrub-level", "aggressive"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)
        # User explicitly set aggressive, should not be overridden to nuclear
        assert options["scrub_level"] == "aggressive"


class TestCLIFlagBehavior:
    """!
    @brief Test CLI flags affect program behavior correctly.
    """

    def test_dry_run_prevents_system_changes(self, monkeypatch, tmp_path) -> None:
        """Test --dry-run passes dry_run=True to scrub execution."""
        monkeypatch.setattr(main, "ensure_admin_and_relaunch_if_needed", _no_op)
        monkeypatch.setattr(main, "enable_vt_mode_if_possible", _no_op)
        monkeypatch.setattr(main, "_resolve_log_directory", lambda c: tmp_path)
        monkeypatch.setattr(main.detect, "gather_office_inventory", lambda **kw: {})
        monkeypatch.setattr(
            main.plan_module,
            "build_plan",
            lambda i, o: [
                {"id": "ctx", "category": "context", "metadata": {"mode": o["mode"]}},
            ],
        )
        monkeypatch.setattr(main.safety, "perform_preflight_checks", lambda p: None)

        scrub_dry_run: list[bool] = []
        monkeypatch.setattr(
            main.scrub,
            "execute_plan",
            lambda plan, dry_run=False, **kw: scrub_dry_run.append(dry_run),
        )
        monkeypatch.setattr(main, "_enforce_runtime_guards", lambda o, *, dry_run=False: None)

        main.main(["--auto-all", "--dry-run", "--logdir", str(tmp_path)])
        assert scrub_dry_run == [True]

    def test_quiet_flag_affects_logging(self, monkeypatch, tmp_path) -> None:
        """Test --quiet flag is parsed and accessible."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--quiet"])
        assert args.quiet is True

    def test_json_flag_affects_logging(self, monkeypatch, tmp_path) -> None:
        """Test --json flag is parsed and accessible."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--json"])
        assert args.json is True

    def test_no_color_flag_accessible(self) -> None:
        """Test --no-color flag is parsed and accessible."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--no-color"])
        assert args.no_color is True

    def test_tui_flags_accessible(self) -> None:
        """Test TUI-related flags are parsed and accessible."""
        parser = main.build_arg_parser()
        args = parser.parse_args(["--tui", "--tui-compact", "--tui-refresh", "100"])
        assert args.tui is True
        assert args.tui_compact is True
        assert args.tui_refresh == 100


class TestConfigFile:
    """Tests for JSON configuration file loading."""

    def test_config_file_loads_options(self, tmp_path) -> None:
        """Test config file options are loaded."""
        config = tmp_path / "config.json"
        config.write_text('{"passes": 5, "dry-run": true, "scrub-level": "aggressive"}')

        parser = main.build_arg_parser()
        # Use --diagnose instead of --auto-all to avoid auto-all's scrub_level override
        args = parser.parse_args(["--diagnose", "--config", str(config)])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)

        assert options["max_passes"] == 5
        assert options["dry_run"] is True
        assert options["scrub_level"] == "aggressive"

    def test_cli_overrides_config_file(self, tmp_path) -> None:
        """Test CLI arguments take precedence over config file."""
        config = tmp_path / "config.json"
        config.write_text('{"passes": 5, "dry-run": true}')

        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--config", str(config), "--passes", "2"])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)

        # CLI --passes=2 should override config passes=5
        assert options["max_passes"] == 2

    def test_config_file_not_found_exits(self, tmp_path) -> None:
        """Test missing config file causes exit."""
        parser = main.build_arg_parser()
        parser.parse_args(["--auto-all", "--config", str(tmp_path / "missing.json")])

        with pytest.raises(SystemExit):
            main._load_config_file(str(tmp_path / "missing.json"))

    def test_invalid_json_exits(self, tmp_path) -> None:
        """Test invalid JSON causes exit."""
        config = tmp_path / "config.json"
        config.write_text("{invalid json")

        with pytest.raises(SystemExit):
            main._load_config_file(str(config))

    def test_non_object_json_exits(self, tmp_path) -> None:
        """Test non-object JSON root causes exit."""
        config = tmp_path / "config.json"
        config.write_text('["array", "not", "object"]')

        with pytest.raises(SystemExit):
            main._load_config_file(str(config))

    def test_no_config_returns_empty_dict(self) -> None:
        """Test None config path returns empty dict."""
        result = main._load_config_file(None)
        assert result == {}

    def test_config_boolean_flags(self, tmp_path) -> None:
        """Test config file boolean flags are properly converted."""
        config = tmp_path / "config.json"
        config.write_text('{"force": true, "clean-msocache": true, "skip-registry": true}')

        parser = main.build_arg_parser()
        args = parser.parse_args(["--auto-all", "--config", str(config)])
        mode = main._determine_mode(args)
        options = main._collect_plan_options(args, mode)

        assert options["force"] is True
        assert options["clean_msocache"] is True
        assert options["skip_registry"] is True


class TestSpinnerUpdateTask:
    """Test spinner update_task preserves elapsed time."""

    def test_update_task_preserves_start_time(self) -> None:
        """Verify update_task doesn't reset the timer."""
        import time

        from office_janitor import spinner

        # Set initial task
        spinner.set_task("Initial task")
        initial_start = spinner._task_start_time

        # Wait a bit
        time.sleep(0.05)

        # Update task (should NOT reset timer)
        spinner.update_task("Updated task")
        updated_start = spinner._task_start_time

        # Start time should be the same
        assert initial_start == updated_start
        assert spinner._current_task == "Updated task"

        # Clean up
        spinner.clear_task()

    def test_set_task_resets_start_time(self) -> None:
        """Verify set_task DOES reset the timer."""
        import time

        from office_janitor import spinner

        # Set initial task
        spinner.set_task("Initial task")
        initial_start = spinner._task_start_time

        # Wait a bit
        time.sleep(0.05)

        # Set new task (SHOULD reset timer)
        spinner.set_task("New task")
        new_start = spinner._task_start_time

        # Start time should be different (later)
        assert new_start > initial_start
        assert spinner._current_task == "New task"

        # Clean up
        spinner.clear_task()

    def test_update_task_starts_timer_if_no_task(self) -> None:
        """Verify update_task starts a timer if no task was active."""
        from office_janitor import spinner

        # Make sure no task is active
        spinner.clear_task()
        assert spinner._current_task is None

        # Update task when none active
        spinner.update_task("New task via update")

        # Should have set start time
        assert spinner._task_start_time > 0
        assert spinner._current_task == "New task via update"

        # Clean up
        spinner.clear_task()
