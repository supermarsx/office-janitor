from __future__ import annotations

import pathlib
import sys
from collections import deque
from collections.abc import Sequence

PROJECT_ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import (  # noqa: E402
    constants,
    exec_utils,
    fs_tools,
    licensing,
    logging_ext,
    processes,
    restore_point,
    tasks_services,
)


def _command_result(
    command: Sequence[str],
    *,
    returncode: int = 0,
    stdout: str = "",
    stderr: str = "",
    skipped: bool = False,
    timed_out: bool = False,
    error: str | None = None,
) -> exec_utils.CommandResult:
    """!
    @brief Helper to fabricate :class:`CommandResult` instances for tests.
    """

    return exec_utils.CommandResult(
        command=[str(part) for part in command],
        returncode=returncode,
        stdout=stdout,
        stderr=stderr,
        duration=0.0,
        skipped=skipped,
        timed_out=timed_out,
        error=error,
    )


def test_cleanup_licenses_runs_commands(monkeypatch, tmp_path) -> None:
    """!
    @brief Licensing cleanup should run PowerShell, OSPP, and filesystem steps.
    """

    logging_ext.setup_logging(tmp_path)
    run_calls: list[list[str]] = []
    removed: list[tuple[list[str], bool]] = []
    exports: list[tuple[list[str], pathlib.Path]] = []

    def fake_run(command, *, event, **kwargs):
        run_calls.append([str(part) for part in command])
        return _command_result(command)

    def fake_remove_paths(paths, *, dry_run: bool):
        removed.append(([str(path) for path in paths], dry_run))

    def fake_export(keys, destination, *, dry_run=False, logger=None):
        exports.append((list(keys), pathlib.Path(destination)))
        return []

    monkeypatch.setattr(licensing.exec_utils, "run_command", fake_run)
    monkeypatch.setattr(licensing.fs_tools, "remove_paths", fake_remove_paths)
    monkeypatch.setattr(licensing.registry_tools, "export_keys", fake_export)

    backup_dir = tmp_path / "backup"
    licensing.cleanup_licenses(
        {
            "paths": [tmp_path / "cache"],
            "uninstall_detected": True,
            "backup_destination": backup_dir,
        }
    )

    assert run_calls, "Expected subprocess commands for licensing cleanup"
    assert run_calls[0][0] == "powershell.exe"
    assert run_calls[1][0] == "cscript.exe"
    assert removed == [([str(tmp_path / "cache")], False)]
    assert exports == [(list(licensing.DEFAULT_REGISTRY_KEYS), backup_dir)]


def test_cleanup_licenses_dry_run(monkeypatch, tmp_path) -> None:
    """!
    @brief Dry-run should avoid invoking external commands.
    """

    logging_ext.setup_logging(tmp_path)
    exported = False
    dry_run_flags: list[bool] = []

    def fake_run(command, *, event, dry_run=False, **kwargs):
        dry_run_flags.append(dry_run)
        return _command_result(command, skipped=dry_run)

    def fake_remove_paths(paths, *, dry_run: bool):
        assert dry_run is True

    def fake_export(*args, **kwargs):
        nonlocal exported
        exported = True

    monkeypatch.setattr(licensing.exec_utils, "run_command", fake_run)
    monkeypatch.setattr(licensing.fs_tools, "remove_paths", fake_remove_paths)
    monkeypatch.setattr(licensing.registry_tools, "export_keys", fake_export)

    licensing.cleanup_licenses(
        {
            "dry_run": True,
            "paths": [tmp_path / "cache"],
            "uninstall_detected": True,
            "backup_destination": tmp_path / "backup",
        }
    )

    assert dry_run_flags and all(dry_run_flags)
    assert not exported


def test_cleanup_licenses_skips_without_uninstall(monkeypatch, tmp_path) -> None:
    """!
    @brief Cleanup should be deferred until uninstall steps finish unless forced.
    """

    logging_ext.setup_logging(tmp_path)
    called = False
    removed = False
    exported = False

    def fake_run(command, *, event, **kwargs):
        nonlocal called
        called = True
        return _command_result(command)

    def fake_remove_paths(paths, *, dry_run: bool):
        nonlocal removed
        removed = True

    def fake_export(*args, **kwargs):
        nonlocal exported
        exported = True

    monkeypatch.setattr(licensing.exec_utils, "run_command", fake_run)
    monkeypatch.setattr(licensing.fs_tools, "remove_paths", fake_remove_paths)
    monkeypatch.setattr(licensing.registry_tools, "export_keys", fake_export)

    licensing.cleanup_licenses({"paths": [tmp_path / "cache"]})

    assert called is False
    assert removed is False
    assert exported is False


def test_render_license_script_defaults_to_constants() -> None:
    """!
    @brief Ensure the embedded PowerShell script references constant values.
    """

    script = licensing._render_license_script({})
    assert constants.LICENSING_GUID_FILTERS["office_family"] in script
    assert constants.LICENSE_DLLS["spp"] in script
    assert constants.OSPP_REGISTRY_PATH in script


def test_remove_paths_deletes_entries(monkeypatch, tmp_path) -> None:
    """!
    @brief Filesystem helper should remove files and directories.
    """

    deleted = tmp_path / "obsolete"
    deleted.mkdir()
    (deleted / "inner.txt").write_text("data", encoding="utf-8")
    file_entry = tmp_path / "file.txt"
    file_entry.write_text("payload", encoding="utf-8")

    calls: list[pathlib.Path] = []
    attrib_calls: list[list[pathlib.Path]] = []

    def fake_reset(path):
        calls.append(path)

    def fake_make(paths, *, dry_run: bool = False):
        attrib_calls.append([pathlib.Path(p) for p in paths])

    monkeypatch.setattr(fs_tools, "make_paths_writable", fake_make)
    monkeypatch.setattr(fs_tools, "reset_acl", fake_reset)
    fs_tools.remove_paths([deleted, file_entry])

    assert not deleted.exists()
    assert not file_entry.exists()
    assert calls == [deleted, file_entry]
    assert attrib_calls


def test_remove_paths_dry_run(monkeypatch, tmp_path) -> None:
    """!
    @brief Dry-run skip should preserve filesystem entries.
    """

    target = tmp_path / "keep.txt"
    target.write_text("keep", encoding="utf-8")

    make_called = False

    def fake_make(paths, *, dry_run: bool = False):
        nonlocal make_called
        make_called = True
        assert dry_run is True

    monkeypatch.setattr(fs_tools, "make_paths_writable", fake_make)
    monkeypatch.setattr(fs_tools, "reset_acl", lambda path: None)
    fs_tools.remove_paths([target], dry_run=True)

    assert target.exists()
    assert make_called is False


def test_reset_acl_invokes_icacls(monkeypatch) -> None:
    """!
    @brief Ensure ``icacls`` is called with expected arguments.
    """

    call_args: list[list[str]] = []

    def fake_run(command, *, event, **kwargs):
        call_args.append([str(part) for part in command])
        return _command_result(command)

    monkeypatch.setattr(fs_tools.exec_utils, "run_command", fake_run)
    fs_tools.reset_acl(pathlib.Path("C:/temp"))

    assert call_args[0][:2] == ["icacls", "C:/temp"]


def test_make_paths_writable_invokes_attrib(monkeypatch, tmp_path) -> None:
    """!
    @brief Attribute clearing should call ``attrib.exe`` for directories and contents.
    """

    commands: list[list[str]] = []

    def fake_run(command, *, event, **kwargs):
        commands.append([str(part) for part in command])
        return _command_result(command)

    target = tmp_path / "dir"
    target.mkdir()

    monkeypatch.setattr(fs_tools.exec_utils, "run_command", fake_run)
    fs_tools.make_paths_writable([target])

    assert commands
    assert commands[0][:2] == ["attrib.exe", "-R"]


def test_terminate_office_processes_invokes_taskkill(monkeypatch, tmp_path) -> None:
    """!
    @brief Process helper should execute ``taskkill`` for each process.
    """

    logging_ext.setup_logging(tmp_path)
    commands: list[list[str]] = []

    def fake_run(command, *, event, **kwargs):
        commands.append([str(part) for part in command])
        return _command_result(command)

    monkeypatch.setattr(processes.exec_utils, "run_command", fake_run)
    processes.terminate_office_processes(["winword.exe", "excel.exe"])

    assert commands
    assert commands[0][0] == "taskkill.exe"
    assert commands[1][0] == "taskkill.exe"


def test_enumerate_processes_filters_patterns(monkeypatch, tmp_path) -> None:
    """!
    @brief Process enumeration should return unique matches per pattern.
    """

    logging_ext.setup_logging(tmp_path)

    def fake_run(command, *, event, **kwargs):
        assert str(command[0]).lower() == "tasklist.exe"
        return _command_result(
            command,
            stdout=(
                """Image Name                     PID Session Name        Session#    Mem Usage\n"""
                "ose.exe 123 Console                    1     12,000 K\n"
                "WINWORD.EXE 456 Console               1     15,000 K\n"
            ),
        )

    monkeypatch.setattr(processes.exec_utils, "run_command", fake_run)

    matches = processes.enumerate_processes(["ose*.exe", "winword.exe", ""])

    assert matches == ["ose.exe", "winword.exe"]


def test_prompt_user_to_close_accepts(monkeypatch, tmp_path) -> None:
    """!
    @brief Prompt should return ``True`` when the operator consents.
    """

    logging_ext.setup_logging(tmp_path)
    answers = iter(["y"])

    result = processes.prompt_user_to_close(
        ["WINWORD.EXE", "excel.exe"], input_func=lambda _: next(answers)
    )

    assert result is True


def test_prompt_user_to_close_declines(monkeypatch, tmp_path) -> None:
    """!
    @brief Prompt should honour refusal after repeated attempts.
    """

    logging_ext.setup_logging(tmp_path)
    answers = iter(["maybe", "n"])

    result = processes.prompt_user_to_close(["WINWORD.EXE"], input_func=lambda _: next(answers))

    assert result is False


def test_prompt_user_to_close_emits_outlook_warning(monkeypatch, tmp_path) -> None:
    """!
    @brief Outlook processes should trigger a reassurance warning and UI event.
    """

    logging_ext.setup_logging(tmp_path)
    event_queue = deque()
    logging_ext.register_ui_event_sink(queue=event_queue)

    answers = iter(["n"])
    processes.prompt_user_to_close(["OUTLOOK.EXE"], input_func=lambda _: next(answers))

    logging_ext.register_ui_event_sink()

    human_logger = logging_ext.get_human_logger()
    for handler in list(human_logger.handlers):
        flush = getattr(handler, "flush", None)
        if callable(flush):
            flush()

    log_dir = logging_ext.get_log_directory()
    assert log_dir is not None
    human_log = log_dir / "human.log"
    assert human_log.exists()
    contents = human_log.read_text(encoding="utf-8")
    assert "OST/PST" in contents
    assert event_queue
    recorded = event_queue[0]
    assert recorded.get("event") == "processes.outlook_reassurance"
    assert "OST/PST" in str(recorded.get("message"))


def test_terminate_process_patterns_uses_enumerator(monkeypatch, tmp_path) -> None:
    """!
    @brief Wildcard termination should rely on :func:`enumerate_processes`.
    """

    logging_ext.setup_logging(tmp_path)
    enumerated: list[list[str]] = []
    killed: list[list[str]] = []

    def fake_enumerate(patterns, *, timeout):
        enumerated.append(list(patterns))
        return ["ose.exe", "integrator.exe"]

    def fake_terminate(names, *, timeout=30):
        killed.append(list(names))

    monkeypatch.setattr(processes, "enumerate_processes", fake_enumerate)
    monkeypatch.setattr(processes, "terminate_office_processes", fake_terminate)

    processes.terminate_process_patterns(["ose*.exe", "integrator.exe"], timeout=15)

    assert enumerated == [["ose*.exe", "integrator.exe"]]
    assert killed == [["ose.exe", "integrator.exe"]]


def test_disable_tasks_respects_dry_run(monkeypatch, tmp_path) -> None:
    """!
    @brief Scheduled task disablement should honour dry-run.
    """

    logging_ext.setup_logging(tmp_path)
    dry_run_flags: list[bool] = []

    def fake_run(command, *, event, dry_run=False, **kwargs):
        dry_run_flags.append(dry_run)
        return _command_result(command, skipped=dry_run)

    monkeypatch.setattr(tasks_services.exec_utils, "run_command", fake_run)
    tasks_services.disable_tasks([r"Microsoft\\Office\\Task"], dry_run=True)

    assert dry_run_flags and all(dry_run_flags)


def test_disable_tasks_executes(monkeypatch, tmp_path) -> None:
    """!
    @brief Without dry-run, ``schtasks`` should be invoked.
    """

    logging_ext.setup_logging(tmp_path)
    commands: list[list[str]] = []

    def fake_run(command, *, event, **kwargs):
        commands.append([str(part) for part in command])
        return _command_result(command)

    monkeypatch.setattr(tasks_services.exec_utils, "run_command", fake_run)
    tasks_services.disable_tasks([r"Microsoft\\Office\\Task"], dry_run=False)

    assert commands
    assert commands[0][0] == "schtasks.exe"


def test_delete_tasks_executes_delete(monkeypatch, tmp_path) -> None:
    """!
    @brief Task deletion should issue ``schtasks /Delete`` commands.
    """

    logging_ext.setup_logging(tmp_path)
    commands: list[list[str]] = []

    def fake_run(command, *, event, **kwargs):
        commands.append([str(part) for part in command])
        return _command_result(command)

    monkeypatch.setattr(tasks_services.exec_utils, "run_command", fake_run)
    tasks_services.delete_tasks([r"Microsoft\\Office\\Cleanup"], dry_run=False)

    assert commands
    assert commands[0][:2] == ["schtasks.exe", "/Delete"]


def test_remove_tasks_aliases_delete(monkeypatch, tmp_path) -> None:
    """!
    @brief Legacy ``remove_tasks`` wrapper should delegate to ``delete_tasks``.
    """

    logging_ext.setup_logging(tmp_path)
    called: list[Sequence[str]] = []

    monkeypatch.setattr(
        tasks_services, "delete_tasks", lambda names, dry_run=False: called.append(list(names))
    )

    tasks_services.remove_tasks(["Task"], dry_run=True)

    assert called == [["Task"]]


def test_stop_services_invokes_sc(monkeypatch, tmp_path) -> None:
    """!
    @brief Service helper should call ``sc.exe`` stop and disable sequences.
    """

    logging_ext.setup_logging(tmp_path)
    commands: list[list[str]] = []

    def fake_run(command, *, event, **kwargs):
        commands.append([str(part) for part in command])
        return _command_result(command)

    monkeypatch.setattr(tasks_services.exec_utils, "run_command", fake_run)
    outcome = tasks_services.stop_services(["ClickToRunSvc"], timeout=10)

    assert len(commands) == 2
    assert commands[0][:2] == ["sc.exe", "stop"]
    assert commands[1][:2] == ["sc.exe", "config"]
    assert outcome == {
        "reboot_required": False,
        "services_requiring_reboot": [],
    }


def test_start_services_invokes_sc(monkeypatch, tmp_path) -> None:
    """!
    @brief Service start helper should invoke ``sc.exe start``.
    """

    logging_ext.setup_logging(tmp_path)
    commands: list[list[str]] = []

    def fake_run(command, *, event, **kwargs):
        commands.append([str(part) for part in command])
        return _command_result(command)

    monkeypatch.setattr(tasks_services.exec_utils, "run_command", fake_run)

    tasks_services.start_services(["ClickToRunSvc"], timeout=5)

    assert commands == [["sc.exe", "start", "ClickToRunSvc"]]


def test_delete_services_executes(monkeypatch, tmp_path) -> None:
    """!
    @brief Service deletion should issue ``sc.exe delete``.
    """

    logging_ext.setup_logging(tmp_path)
    commands: list[list[str]] = []

    def fake_run(command, *, event, **kwargs):
        commands.append([str(part) for part in command])
        return _command_result(command)

    monkeypatch.setattr(tasks_services.exec_utils, "run_command", fake_run)
    tasks_services.delete_services(["ose"], dry_run=False)

    assert commands
    assert commands[0][:2] == ["sc.exe", "delete"]


def test_query_service_status_retries_and_parses(monkeypatch, tmp_path) -> None:
    """!
    @brief Status query should retry on timeouts and parse the ``STATE`` line.
    """

    logging_ext.setup_logging(tmp_path)
    attempts: list[list[str]] = []

    def fake_run(command, *, event, **kwargs):
        attempts.append([str(part) for part in command])
        if len(attempts) < 3:
            return _command_result(
                command,
                returncode=1,
                stdout="",
                stderr="",
                timed_out=True,
                error="timeout",
            )
        return _command_result(
            command,
            stdout="""\nSERVICE_NAME: ClickToRunSvc\n        STATE              : 4  RUNNING\n""",
        )

    monkeypatch.setattr(tasks_services.exec_utils, "run_command", fake_run)
    monkeypatch.setattr(tasks_services.time, "sleep", lambda seconds: None)

    status = tasks_services.query_service_status("ClickToRunSvc", retries=3, delay=0.1, timeout=1)

    assert len(attempts) == 3
    assert status == "RUNNING"


def test_create_restore_point_uses_powershell(monkeypatch, tmp_path) -> None:
    """!
    @brief Restore point helper should call ``powershell.exe``.
    """

    logging_ext.setup_logging(tmp_path)
    captured: list[list[str]] = []

    def fake_run(command, *, event, **kwargs):
        captured.append([str(part) for part in command])
        return _command_result(command)

    monkeypatch.setattr(restore_point.os, "name", "nt")
    monkeypatch.setattr(restore_point.exec_utils, "run_command", fake_run)
    result = restore_point.create_restore_point("Before Office cleanup")

    assert captured
    assert captured[0][0] == "powershell.exe"
    assert "SystemRestore" in captured[0][-1]
    assert result is True


def test_create_restore_point_dry_run_skips_execution(monkeypatch, tmp_path) -> None:
    """!
    @brief Dry-run mode should avoid invoking PowerShell.
    """

    logging_ext.setup_logging(tmp_path)
    calls: list[list[str]] = []

    def fake_run(command, *, event, dry_run=False, **kwargs):
        calls.append((list(command), dry_run))
        return _command_result(command, skipped=dry_run)

    monkeypatch.setattr(restore_point.os, "name", "nt")
    monkeypatch.setattr(restore_point.exec_utils, "run_command", fake_run)

    result = restore_point.create_restore_point("Simulated cleanup", dry_run=True)

    assert result is True
    assert calls and all(flag for _, flag in calls)
