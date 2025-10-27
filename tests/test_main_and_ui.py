"""Integration tests for CLI and UI layers."""
from __future__ import annotations

import json
import pathlib
import sys
from typing import List

import pytest

PROJECT_ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import main, ui, version  # noqa: E402
from office_janitor import tui as tui_module  # noqa: E402


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
    monkeypatch.setattr(main.detect, "gather_office_inventory", lambda: inventory)

    recorded: List[tuple[str, object]] = []

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
    monkeypatch.setattr(main.safety, "perform_preflight_checks", lambda plan: recorded.append(("safety", len(plan))))

    scrub_calls: List[bool] = []
    monkeypatch.setattr(main.scrub, "execute_plan", lambda plan, dry_run=False: scrub_calls.append(bool(dry_run)))

    exit_code = main.main(["--auto-all", "--dry-run", "--logdir", str(tmp_path / "logs")])

    assert exit_code == 0
    assert scrub_calls == [True]
    assert recorded[0][0] == "plan"
    assert recorded[0][1]["mode"] == "auto-all"
    assert any(item[0] == "safety" for item in recorded)


def test_main_diagnose_skips_execution(monkeypatch, tmp_path) -> None:
    """!
    @brief Diagnostics mode must avoid executing the scrubber.
    """

    monkeypatch.setattr(main, "ensure_admin_and_relaunch_if_needed", _no_op)
    monkeypatch.setattr(main, "enable_vt_mode_if_possible", _no_op)
    monkeypatch.setattr(main, "_resolve_log_directory", lambda candidate: tmp_path)

    monkeypatch.setattr(main.detect, "gather_office_inventory", lambda: {"msi": [], "c2r": [], "filesystem": []})

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

    scrub_calls: List[bool] = []
    monkeypatch.setattr(main.scrub, "execute_plan", lambda plan, dry_run=False: scrub_calls.append(True))

    exit_code = main.main(["--diagnose", "--logdir", str(tmp_path / "logs")])

    assert exit_code == 0
    assert scrub_calls == []


def test_main_interactive_uses_cli(monkeypatch, tmp_path) -> None:
    """!
    @brief Without mode flags, the plain menu should launch.
    """

    monkeypatch.setattr(main, "ensure_admin_and_relaunch_if_needed", _no_op)
    monkeypatch.setattr(main, "enable_vt_mode_if_possible", _no_op)
    monkeypatch.setattr(main, "_resolve_log_directory", lambda candidate: tmp_path)
    monkeypatch.setattr(main, "_should_use_tui", lambda args: False)

    captured = {}

    def fake_run_cli(app_state):  # type: ignore[no-untyped-def]
        captured["app_state"] = app_state

    monkeypatch.setattr(main.ui, "run_cli", fake_run_cli)
    monkeypatch.setattr(main.tui, "run_tui", lambda app_state: (_ for _ in ()).throw(AssertionError("TUI not expected")))

    exit_code = main.main(["--logdir", str(tmp_path / "logs")])

    assert exit_code == 0
    assert "detector" in captured["app_state"]


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
    monkeypatch.setattr(main.detect, "gather_office_inventory", lambda: inventory)

    captured_options: dict = {}

    def fake_plan(inv, options):  # type: ignore[no-untyped-def]
        captured_options.update(options)
        return [
            {"id": "context", "category": "context", "metadata": {"mode": options.get("mode"), "options": dict(options)}},
            {"id": "filesystem", "category": "filesystem-cleanup", "metadata": {"paths": ["C:/Office"]}},
        ]

    monkeypatch.setattr(main.plan_module, "build_plan", fake_plan)
    monkeypatch.setattr(main.safety, "perform_preflight_checks", lambda plan: None)

    scrub_calls: List[bool] = []

    def fake_execute(plan, dry_run=False):  # type: ignore[no-untyped-def]
        scrub_calls.append(bool(dry_run))

    monkeypatch.setattr(main.scrub, "execute_plan", fake_execute)

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


def test_ui_run_cli_detect_option(monkeypatch) -> None:
    """!
    @brief Menu option 1 should call the detector and exit cleanly.
    """

    events: List[str] = []
    inputs = iter(["1", "7"])

    def fake_input(prompt: str) -> str:
        return next(inputs)

    app_state = {
        "args": type("Args", (), {"quiet": False, "dry_run": False, "no_restore_point": False, "logdir": "logs", "backup": None})(),
        "detector": lambda: events.append("detect") or {"msi": [1], "c2r": [], "filesystem": []},
        "planner": lambda inventory, overrides=None: (_ for _ in ()).throw(AssertionError("planner not expected")),
        "executor": lambda plan, overrides=None: (_ for _ in ()).throw(AssertionError("executor not expected")),
        "input": fake_input,
    }

    ui.run_cli(app_state)

    assert events == ["detect"]


def test_ui_run_cli_auto_all_executes(monkeypatch) -> None:
    """!
    @brief Menu option 2 should plan and execute using overrides.
    """

    events: List[tuple[str, object]] = []
    inputs = iter(["2", "7"])

    def fake_input(prompt: str) -> str:
        return next(inputs)

    def fake_detector():
        events.append(("detect", None))
        return {"msi": [], "c2r": [], "filesystem": []}

    def fake_planner(inventory, overrides=None):
        events.append(("plan", overrides))
        return [
            {"id": "context", "category": "context", "metadata": {"mode": overrides.get("mode") if overrides else "interactive", "dry_run": False, "target_versions": [], "unsupported_targets": [], "options": {}}},
            {"id": "registry-0", "category": "registry-cleanup", "metadata": {"keys": []}},
        ]

    def fake_executor(plan, overrides=None):
        events.append(("execute", overrides))

    app_state = {
        "args": type("Args", (), {"quiet": False, "dry_run": False, "no_restore_point": False, "logdir": "logs", "backup": None})(),
        "detector": fake_detector,
        "planner": fake_planner,
        "executor": fake_executor,
        "input": fake_input,
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

    events: List[tuple[str, object]] = []
    inputs = iter(["3", "2016,365", "visio,project", "7"])

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
        "args": type("Args", (), {"quiet": False, "dry_run": False, "no_restore_point": False, "logdir": "logs", "backup": None})(),
        "detector": fake_detector,
        "planner": fake_planner,
        "executor": fake_executor,
        "input": fake_input,
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

    events: List[str] = []

    app_state = {
        "args": type("Args", (), {"quiet": False, "json": True, "dry_run": False, "no_restore_point": False, "logdir": "logs", "backup": None})(),
        "detector": lambda: events.append("detect"),
        "planner": lambda inventory, overrides=None: events.append("plan"),
        "executor": lambda plan, overrides=None: events.append("execute"),
    }

    ui.run_cli(app_state)

    assert events == []


def test_tui_falls_back_without_ansi(monkeypatch) -> None:
    """!
    @brief When ANSI support is missing the TUI should delegate to the CLI.
    """

    monkeypatch.setattr(tui_module, "_supports_ansi", lambda stream=None: False)

    invoked: List[str] = []

    def fake_run_cli(app_state):  # type: ignore[no-untyped-def]
        invoked.append("cli")

    monkeypatch.setattr(ui, "run_cli", fake_run_cli)

    tui_module.run_tui({
        "args": type("Args", (), {"no_color": False, "quiet": False})(),
        "detector": lambda: {},
        "planner": lambda inventory, overrides=None: [],
        "executor": lambda plan, overrides=None: None,
    })

    assert invoked == ["cli"]


def test_tui_commands_drive_backends(monkeypatch) -> None:
    """!
    @brief Key commands should call detector, planner, and executor in order.
    """

    monkeypatch.setattr(tui_module, "_supports_ansi", lambda stream=None: True)
    monkeypatch.setattr(tui_module, "_spinner", lambda duration, message: None)

    keys = iter(["d", "p", "r", "q"])

    def reader() -> str:
        return next(keys)

    events: List[str] = []

    def fake_detector():
        events.append("detect")
        return {"msi": [], "c2r": [], "filesystem": []}

    def fake_planner(inventory, overrides=None):
        events.append("plan")
        return [
            {"id": "context", "category": "context", "metadata": {"mode": "interactive", "dry_run": False, "target_versions": [], "unsupported_targets": [], "options": {}}},
            {"id": "filesystem-0", "category": "filesystem-cleanup", "metadata": {"paths": []}},
        ]

    def fake_executor(plan, overrides=None):
        events.append("execute")

    tui_module.run_tui({
        "args": type("Args", (), {"no_color": False, "quiet": False})(),
        "detector": fake_detector,
        "planner": fake_planner,
        "executor": fake_executor,
        "key_reader": reader,
    })

    assert events == ["detect", "plan", "execute"]


def test_tui_respects_quiet_flag(monkeypatch) -> None:
    """!
    @brief Quiet mode should prevent the TUI from starting an event loop.
    """

    monkeypatch.setattr(tui_module, "_supports_ansi", lambda stream=None: True)

    invoked: List[str] = []

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
        }
    )

    assert invoked == []


def test_tui_auto_mode_invokes_overrides(monkeypatch) -> None:
    """!
    @brief The ``A`` command should trigger auto scrub with overrides.
    """

    monkeypatch.setattr(tui_module, "_supports_ansi", lambda stream=None: True)
    monkeypatch.setattr(tui_module, "_spinner", lambda duration, message: None)

    keys = iter(["a", "q"])

    def reader() -> str:
        return next(keys)

    events: List[tuple[str, object]] = []

    def fake_detector():
        events.append(("detect", None))
        return {"msi": [], "c2r": [], "filesystem": []}

    def fake_planner(inventory, overrides=None):
        events.append(("plan", overrides))
        return [
            {"id": "context", "category": "context", "metadata": {"mode": overrides.get("mode"), "options": dict(overrides)}}
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

    keys = iter(["t", "q"])

    def reader() -> str:
        return next(keys)

    prompts = iter(["2016,365", "visio"])

    monkeypatch.setattr(tui_module, "_read_input_line", lambda prompt: next(prompts))

    events: List[tuple[str, object]] = []

    def fake_detector():
        events.append(("detect", None))
        return {"msi": [], "c2r": [], "filesystem": []}

    def fake_planner(inventory, overrides=None):
        events.append(("plan", overrides))
        return [
            {"id": "context", "category": "context", "metadata": {"mode": overrides.get("mode"), "options": dict(overrides)}}
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

    monkeypatch.setattr(main.detect, "gather_office_inventory", lambda: {"msi": ["Office"], "c2r": [], "filesystem": []})

    def fake_plan(inv, options):  # type: ignore[no-untyped-def]
        return [
            {"id": "context", "category": "context", "metadata": {"mode": options.get("mode"), "options": dict(options)}},
            {"id": "filesystem", "category": "filesystem-cleanup", "metadata": {"paths": ["C:/Office"]}},
        ]

    monkeypatch.setattr(main.plan_module, "build_plan", fake_plan)
    monkeypatch.setattr(main.safety, "perform_preflight_checks", lambda plan: None)
    monkeypatch.setattr(main.scrub, "execute_plan", lambda plan, dry_run=False: None)

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

