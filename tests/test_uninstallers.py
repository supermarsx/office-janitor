"""!
@brief Validate uninstall helper command composition and retry behaviour.
@details Ensures :mod:`msi_uninstall` and :mod:`c2r_uninstall` build the
expected ``msiexec``/Click-to-Run commands, honour dry-run semantics, and
surface failures through informative exceptions.
"""

from __future__ import annotations

import pathlib
import sys
from typing import List, Tuple

import pytest

PROJECT_ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import c2r_uninstall, command_runner, logging_ext, msi_uninstall


def _command_result(command: List[str], returncode: int = 0, *, skipped: bool = False) -> command_runner.CommandResult:
    """!
    @brief Convenience factory for :class:`CommandResult` instances.
    """

    return command_runner.CommandResult(
        command=command,
        returncode=returncode,
        stdout="",
        stderr="",
        duration=0.1,
        skipped=skipped,
    )


def test_msi_uninstall_builds_msiexec_command(monkeypatch, tmp_path) -> None:
    """!
    @brief Ensure ``msiexec`` commands are constructed and verification runs.
    """

    logging_ext.setup_logging(tmp_path)
    executed: List[List[str]] = []
    state = {"present": True}

    def fake_run_command(
        command: List[str],
        *,
        event: str,
        timeout: float | None = None,
        dry_run: bool = False,
        human_message: str | None = None,
        extra: dict | None = None,
    ) -> command_runner.CommandResult:
        executed.append(command)
        state["present"] = False
        return _command_result(command)

    monkeypatch.setattr(msi_uninstall.command_runner, "run_command", fake_run_command)
    monkeypatch.setattr(msi_uninstall.registry_tools, "key_exists", lambda *_, **__: state["present"])
    monkeypatch.setattr(msi_uninstall.time, "sleep", lambda *_: None)

    record = {
        "product": "Office",
        "product_code": "{91160000-0011-0000-0000-0000000FF1CE}",
        "uninstall_handles": [
            "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{91160000-0011-0000-0000-0000000FF1CE}"
        ],
    }
    msi_uninstall.uninstall_products([record])

    assert executed, "Expected msiexec to be invoked"
    command = executed[0]
    assert command[0].lower().endswith("msiexec.exe")
    assert command[1] == "/x"
    assert command[2].startswith("{")
    assert "/qb!" in command
    assert "/norestart" in command


def test_msi_uninstall_dry_run(monkeypatch, tmp_path) -> None:
    """!
    @brief Dry-run mode should record the plan without executing ``msiexec``.
    """

    logging_ext.setup_logging(tmp_path)
    called = False

    def fake_run_command(
        command: List[str],
        *,
        event: str,
        timeout: float | None = None,
        dry_run: bool = False,
        human_message: str | None = None,
        extra: dict | None = None,
    ) -> command_runner.CommandResult:
        nonlocal called
        called = True
        assert dry_run is True
        return _command_result(command, skipped=True)

    monkeypatch.setattr(msi_uninstall.command_runner, "run_command", fake_run_command)
    monkeypatch.setattr(msi_uninstall.registry_tools, "key_exists", lambda *_, **__: True)

    msi_uninstall.uninstall_products(
        [
            {
                "product": "Office",
                "product_code": "{91160000-0011-0000-0000-0000000FF1CE}",
            }
        ],
        dry_run=True,
    )

    assert called, "Dry-run should still build the command"


def test_msi_uninstall_reports_failure(monkeypatch, tmp_path) -> None:
    """!
    @brief Non-zero return codes propagate as ``RuntimeError`` instances.
    """

    logging_ext.setup_logging(tmp_path)

    def fake_run_command(
        command: List[str],
        *,
        event: str,
        timeout: float | None = None,
        dry_run: bool = False,
        human_message: str | None = None,
        extra: dict | None = None,
    ) -> command_runner.CommandResult:
        return _command_result(command, returncode=1603)

    monkeypatch.setattr(msi_uninstall.command_runner, "run_command", fake_run_command)
    monkeypatch.setattr(msi_uninstall.registry_tools, "key_exists", lambda *_, **__: True)

    with pytest.raises(RuntimeError) as excinfo:
        msi_uninstall.uninstall_products(["{BAD-CODE}"])

    assert "BAD-CODE" in str(excinfo.value)


def test_c2r_uninstall_prefers_client(monkeypatch, tmp_path) -> None:
    """!
    @brief OfficeC2RClient.exe should be preferred when available.
    """

    logging_ext.setup_logging(tmp_path)
    executed: List[Tuple[List[str], dict]] = []
    state = {"present": True}

    def fake_run_command(
        command: List[str],
        *,
        event: str,
        timeout: float | None = None,
        dry_run: bool = False,
        human_message: str | None = None,
        extra: dict | None = None,
    ) -> command_runner.CommandResult:
        executed.append((command, extra or {}))
        state["present"] = False
        return _command_result(command)

    monkeypatch.setattr(c2r_uninstall.command_runner, "run_command", fake_run_command)
    monkeypatch.setattr(c2r_uninstall.tasks_services, "stop_services", lambda services, *, timeout=30: None)
    monkeypatch.setattr(c2r_uninstall, "_handles_present", lambda target: state["present"])
    monkeypatch.setattr(c2r_uninstall, "_install_paths_present", lambda target: state["present"])
    monkeypatch.setattr(c2r_uninstall.time, "sleep", lambda *_: None)

    client_path = tmp_path / "OfficeC2RClient.exe"
    client_path.write_text("")

    config = {
        "product": "Microsoft 365 Apps",
        "release_ids": ["O365ProPlusRetail"],
        "client_paths": [client_path],
        "uninstall_handles": [
            "HKLM\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\ProductReleaseIDs\\O365ProPlusRetail"
        ],
        "install_path": str(tmp_path),
    }

    c2r_uninstall.uninstall_products(config)

    assert executed, "Expected OfficeC2RClient.exe invocation"
    command, metadata = executed[0]
    assert command[0].endswith("OfficeC2RClient.exe")
    for arg in c2r_uninstall.C2R_CLIENT_ARGS:
        assert arg in command
    assert metadata.get("executable", "").endswith("OfficeC2RClient.exe")


def test_c2r_uninstall_fallback_to_setup(monkeypatch, tmp_path) -> None:
    """!
    @brief ``setup.exe`` fallback should run when the client is missing.
    """

    logging_ext.setup_logging(tmp_path)
    executed: List[List[str]] = []
    state = {"present": True}

    def fake_run_command(
        command: List[str],
        *,
        event: str,
        timeout: float | None = None,
        dry_run: bool = False,
        human_message: str | None = None,
        extra: dict | None = None,
    ) -> command_runner.CommandResult:
        executed.append(command)
        state["present"] = False
        return _command_result(command)

    monkeypatch.setattr(c2r_uninstall.command_runner, "run_command", fake_run_command)
    monkeypatch.setattr(c2r_uninstall.tasks_services, "stop_services", lambda services, *, timeout=30: None)
    monkeypatch.setattr(c2r_uninstall, "_handles_present", lambda target: state["present"])
    monkeypatch.setattr(c2r_uninstall, "_install_paths_present", lambda target: state["present"])
    monkeypatch.setattr(c2r_uninstall.time, "sleep", lambda *_: None)

    setup_path = tmp_path / "setup.exe"
    setup_path.write_text("")

    config = {
        "release_ids": ["O365ProPlusRetail", "VisioProRetail"],
        "setup_paths": [setup_path],
        "install_path": str(tmp_path),
    }

    c2r_uninstall.uninstall_products(config)

    assert executed, "Expected setup.exe invocation"
    assert all(cmd[0].endswith("setup.exe") for cmd in executed)
    assert {cmd[2] for cmd in executed} == {"O365ProPlusRetail", "VisioProRetail"}


def test_c2r_uninstall_dry_run(monkeypatch, tmp_path) -> None:
    """!
    @brief Dry-run should not stop services or verify removal.
    """

    logging_ext.setup_logging(tmp_path)
    executed: List[List[str]] = []

    def fake_run_command(
        command: List[str],
        *,
        event: str,
        timeout: float | None = None,
        dry_run: bool = False,
        human_message: str | None = None,
        extra: dict | None = None,
    ) -> command_runner.CommandResult:
        executed.append(command)
        return _command_result(command, skipped=True)

    monkeypatch.setattr(c2r_uninstall.command_runner, "run_command", fake_run_command)

    def fail_if_called(*args, **kwargs):
        raise AssertionError("Services should not stop in dry-run")

    monkeypatch.setattr(c2r_uninstall.tasks_services, "stop_services", fail_if_called)
    monkeypatch.setattr(c2r_uninstall.registry_tools, "key_exists", lambda *_, **__: True)

    client_path = tmp_path / "OfficeC2RClient.exe"
    client_path.write_text("")

    c2r_uninstall.uninstall_products(
        {
            "release_ids": ["O365ProPlusRetail"],
            "client_paths": [client_path],
        },
        dry_run=True,
    )

    assert executed, "Dry-run should still build a command"
