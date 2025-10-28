"""!
@brief Exec utils behaviour tests.
@details Validates environment sanitisation, dry-run flows, and subprocess
logging behaviour for :mod:`office_janitor.exec_utils`.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path
from types import SimpleNamespace
from typing import Dict, List

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import exec_utils  # noqa: E402


class _StubLogger:
    """!
    @brief Lightweight logger capturing structured log calls.
    """

    def __init__(self) -> None:
        self.records: List[tuple[str, str, Dict[str, object]]] = []

    def _record(self, level: str, message: str, args: tuple[object, ...], kwargs: Dict[str, object]) -> None:
        text = message % args if args else message
        self.records.append((level, text, dict(kwargs)))

    def info(self, message: str, *args: object, **kwargs: object) -> None:  # noqa: D401 - logging compatibility
        self._record("info", message, args, kwargs)

    def warning(self, message: str, *args: object, **kwargs: object) -> None:  # noqa: D401 - logging compatibility
        self._record("warning", message, args, kwargs)

    def error(self, message: str, *args: object, **kwargs: object) -> None:  # noqa: D401 - logging compatibility
        self._record("error", message, args, kwargs)


def test_sanitize_environment_strips_blocklist_and_overrides() -> None:
    """!
    @brief Ensure sanitisation removes Python-specific variables and applies overrides.
    """

    base_env = {"PYTHONPATH": "should_remove", "KEEP": "1", "LANG": "C"}

    sanitized = exec_utils.sanitize_environment(
        base_env=base_env,
        inherit=False,
        extra={"NEW": "value"},
        remove=["KEEP"],
    )

    assert "PYTHONPATH" not in sanitized
    assert "KEEP" not in sanitized
    assert sanitized["LANG"] == "C"
    assert sanitized["NEW"] == "value"


def test_run_command_dry_run_logs_without_invocation(monkeypatch: pytest.MonkeyPatch) -> None:
    """!
    @brief Dry-run execution should skip subprocess invocation while logging intent.
    """

    human_logger = _StubLogger()
    machine_logger = _StubLogger()

    monkeypatch.setattr(exec_utils.logging_ext, "get_human_logger", lambda: human_logger)
    monkeypatch.setattr(exec_utils.logging_ext, "get_machine_logger", lambda: machine_logger)

    calls: List[List[str]] = []

    def fake_run(*args: object, **kwargs: object) -> None:  # pragma: no cover - should not be called
        calls.append(list(args[0]))
        raise AssertionError("subprocess.run should not be invoked in dry-run mode")

    monkeypatch.setattr(exec_utils.subprocess, "run", fake_run)

    result = exec_utils.run_command(
        ["powershell", "-Command", "Write-Host"],
        event="sample",
        dry_run=True,
        human_message="Executing sample",
    )

    assert result.skipped is True
    assert result.returncode == 0
    assert not calls
    assert machine_logger.records[0][1] == "sample_plan"
    assert machine_logger.records[0][2]["extra"]["command"] == ["powershell", "-Command", "Write-Host"]
    assert machine_logger.records[1][1] == "sample_dry_run"
    assert machine_logger.records[1][2]["extra"]["command"] == ["powershell", "-Command", "Write-Host"]
    assert "[dry-run]" in human_logger.records[0][1]


def test_run_command_executes_with_sanitized_environment(monkeypatch: pytest.MonkeyPatch) -> None:
    """!
    @brief Execute path should sanitise the environment before invoking subprocesses.
    """

    human_logger = _StubLogger()
    machine_logger = _StubLogger()

    monkeypatch.setattr(exec_utils.logging_ext, "get_human_logger", lambda: human_logger)
    monkeypatch.setattr(exec_utils.logging_ext, "get_machine_logger", lambda: machine_logger)

    captured_env: Dict[str, str] = {}

    def fake_run(command, *, capture_output, text, timeout, check, env, cwd):
        captured_env.update(env)
        return SimpleNamespace(returncode=0, stdout="ok", stderr="")

    monkeypatch.setattr(exec_utils.subprocess, "run", fake_run)

    result = exec_utils.run_command(
        ["cmd"],
        event="sanity",
        env={"PYTHONPATH": "value", "KEEP": "1"},
        inherit_env=False,
        env_overrides={"EXTRA": "2"},
        env_remove=["KEEP"],
    )

    assert result.stdout == "ok"
    assert captured_env.get("PYTHONPATH") is None
    assert captured_env.get("KEEP") is None
    assert captured_env["EXTRA"] == "2"
    assert machine_logger.records[-1][1] == "sanity_result"
    assert machine_logger.records[-1][2]["extra"]["return_code"] == 0


def test_run_command_check_raises_on_failure(monkeypatch: pytest.MonkeyPatch) -> None:
    """!
    @brief ``check=True`` should surface non-zero exit codes as ``CalledProcessError``.
    """

    human_logger = _StubLogger()
    machine_logger = _StubLogger()

    monkeypatch.setattr(exec_utils.logging_ext, "get_human_logger", lambda: human_logger)
    monkeypatch.setattr(exec_utils.logging_ext, "get_machine_logger", lambda: machine_logger)

    def fake_run(command, *, capture_output, text, timeout, check, env, cwd):
        return SimpleNamespace(returncode=5, stdout="", stderr="boom")

    monkeypatch.setattr(exec_utils.subprocess, "run", fake_run)

    with pytest.raises(subprocess.CalledProcessError) as excinfo:
        exec_utils.run_command(["cmd"], event="failure", check=True)

    assert excinfo.value.returncode == 5
    assert machine_logger.records[-1][1] == "failure_result"
    assert machine_logger.records[-1][2]["extra"]["return_code"] == 5
    warning_messages = [record for record in human_logger.records if record[0] == "warning"]
    assert warning_messages and "exited with" in warning_messages[0][1]
