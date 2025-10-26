from __future__ import annotations

import pathlib
import sys
from typing import List

PROJECT_ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import (
    constants,
    fs_tools,
    licensing,
    logging_ext,
    processes,
    restore_point,
    tasks_services,
)


class _Result:
    """!
    @brief Minimal mock of :class:`subprocess.CompletedProcess` used for testing.
    """

    def __init__(self, returncode: int = 0, stdout: str = "", stderr: str = "") -> None:
        self.returncode = returncode
        self.stdout = stdout
        self.stderr = stderr


def test_cleanup_licenses_runs_commands(monkeypatch, tmp_path) -> None:
    """!
    @brief Licensing cleanup should run PowerShell, OSPP, and filesystem steps.
    """

    logging_ext.setup_logging(tmp_path)
    run_calls: List[List[str]] = []
    removed: List[tuple[List[str], bool]] = []

    def fake_run(cmd, **kwargs):
        run_calls.append(cmd)
        return _Result()

    def fake_remove_paths(paths, *, dry_run: bool):
        removed.append(([str(path) for path in paths], dry_run))

    monkeypatch.setattr(licensing.subprocess, "run", fake_run)
    monkeypatch.setattr(licensing.fs_tools, "remove_paths", fake_remove_paths)

    licensing.cleanup_licenses({"paths": [tmp_path / "cache"]})

    assert run_calls, "Expected subprocess commands for licensing cleanup"
    assert run_calls[0][0] == "powershell.exe"
    assert run_calls[1][0] == "cscript.exe"
    assert removed == [([str(tmp_path / "cache")], False)]


def test_cleanup_licenses_dry_run(monkeypatch, tmp_path) -> None:
    """!
    @brief Dry-run should avoid invoking external commands.
    """

    logging_ext.setup_logging(tmp_path)
    called = False

    def fake_run(cmd, **kwargs):
        nonlocal called
        called = True
        return _Result()

    def fake_remove_paths(paths, *, dry_run: bool):
        assert dry_run is True

    monkeypatch.setattr(licensing.subprocess, "run", fake_run)
    monkeypatch.setattr(licensing.fs_tools, "remove_paths", fake_remove_paths)

    licensing.cleanup_licenses({"dry_run": True, "paths": [tmp_path / "cache"]})

    assert not called


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

    calls: List[pathlib.Path] = []
    attrib_calls: List[List[pathlib.Path]] = []

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

    call_args: List[List[str]] = []

    def fake_run(cmd, **kwargs):
        call_args.append(cmd)
        return _Result()

    monkeypatch.setattr(fs_tools.subprocess, "run", fake_run)
    fs_tools.reset_acl(pathlib.Path("C:/temp"))

    assert call_args[0][:2] == ["icacls", "C:/temp"]


def test_make_paths_writable_invokes_attrib(monkeypatch, tmp_path) -> None:
    """!
    @brief Attribute clearing should call ``attrib.exe`` for directories and contents.
    """

    commands: List[List[str]] = []

    def fake_run(cmd, **kwargs):
        commands.append(cmd)
        return _Result()

    target = tmp_path / "dir"
    target.mkdir()

    monkeypatch.setattr(fs_tools.subprocess, "run", fake_run)
    fs_tools.make_paths_writable([target])

    assert commands
    assert commands[0][:2] == ["attrib.exe", "-R"]


def test_terminate_office_processes_invokes_taskkill(monkeypatch, tmp_path) -> None:
    """!
    @brief Process helper should execute ``taskkill`` for each process.
    """

    logging_ext.setup_logging(tmp_path)
    commands: List[List[str]] = []

    def fake_run(cmd, **kwargs):
        commands.append(cmd)
        return _Result()

    monkeypatch.setattr(processes.subprocess, "run", fake_run)
    processes.terminate_office_processes(["winword.exe", "excel.exe"])

    assert commands
    assert commands[0][0] == "taskkill.exe"
    assert commands[1][0] == "taskkill.exe"


def test_terminate_process_patterns_uses_tasklist(monkeypatch, tmp_path) -> None:
    """!
    @brief Wildcard termination should enumerate processes before invoking taskkill.
    """

    logging_ext.setup_logging(tmp_path)
    enumerations: List[List[str]] = []
    killed: List[List[str]] = []

    def fake_run(cmd, **kwargs):
        if cmd[0].lower() == "tasklist.exe":
            enumerations.append(cmd)
            return _Result(
                stdout="""Image Name                     PID Session Name        Session#    Mem Usage\nose.exe 123 Console                    1     12,000 K\nIntegrator.exe 456 Console                    1     15,000 K\n""",
            )
        raise AssertionError("Unexpected command")

    def fake_terminate(names, *, timeout=30):
        killed.append(list(names))

    monkeypatch.setattr(processes.subprocess, "run", fake_run)
    monkeypatch.setattr(processes, "terminate_office_processes", fake_terminate)

    processes.terminate_process_patterns(["ose*.exe", "integrator.exe"])

    assert enumerations
    assert killed == [["ose.exe", "integrator.exe"]]


def test_disable_tasks_respects_dry_run(monkeypatch, tmp_path) -> None:
    """!
    @brief Scheduled task disablement should honour dry-run.
    """

    logging_ext.setup_logging(tmp_path)
    called = False

    def fake_run(cmd, **kwargs):
        nonlocal called
        called = True
        return _Result()

    monkeypatch.setattr(tasks_services.subprocess, "run", fake_run)
    tasks_services.disable_tasks([r"Microsoft\\Office\\Task"], dry_run=True)

    assert not called


def test_disable_tasks_executes(monkeypatch, tmp_path) -> None:
    """!
    @brief Without dry-run, ``schtasks`` should be invoked.
    """

    logging_ext.setup_logging(tmp_path)
    commands: List[List[str]] = []

    def fake_run(cmd, **kwargs):
        commands.append(cmd)
        return _Result()

    monkeypatch.setattr(tasks_services.subprocess, "run", fake_run)
    tasks_services.disable_tasks([r"Microsoft\\Office\\Task"], dry_run=False)

    assert commands
    assert commands[0][0] == "schtasks.exe"


def test_remove_tasks_executes_delete(monkeypatch, tmp_path) -> None:
    """!
    @brief Task deletion should issue ``schtasks /Delete`` commands.
    """

    logging_ext.setup_logging(tmp_path)
    commands: List[List[str]] = []

    def fake_run(cmd, **kwargs):
        commands.append(cmd)
        return _Result()

    monkeypatch.setattr(tasks_services.subprocess, "run", fake_run)
    tasks_services.remove_tasks([r"Microsoft\\Office\\Cleanup"], dry_run=False)

    assert commands
    assert commands[0][:2] == ["schtasks.exe", "/Delete"]


def test_stop_services_invokes_sc(monkeypatch, tmp_path) -> None:
    """!
    @brief Service helper should call ``sc.exe`` stop and disable sequences.
    """

    logging_ext.setup_logging(tmp_path)
    commands: List[List[str]] = []

    def fake_run(cmd, **kwargs):
        commands.append(cmd)
        return _Result()

    monkeypatch.setattr(tasks_services.subprocess, "run", fake_run)
    tasks_services.stop_services(["ClickToRunSvc"], timeout=10)

    assert len(commands) == 2
    assert commands[0][:2] == ["sc.exe", "stop"]
    assert commands[1][:2] == ["sc.exe", "config"]


def test_delete_services_executes(monkeypatch, tmp_path) -> None:
    """!
    @brief Service deletion should issue ``sc.exe delete``.
    """

    logging_ext.setup_logging(tmp_path)
    commands: List[List[str]] = []

    def fake_run(cmd, **kwargs):
        commands.append(cmd)
        return _Result()

    monkeypatch.setattr(tasks_services.subprocess, "run", fake_run)
    tasks_services.delete_services(["ose"], dry_run=False)

    assert commands
    assert commands[0][:2] == ["sc.exe", "delete"]


def test_create_restore_point_uses_powershell(monkeypatch, tmp_path) -> None:
    """!
    @brief Restore point helper should call ``powershell.exe``.
    """

    logging_ext.setup_logging(tmp_path)
    captured: List[List[str]] = []

    def fake_run(cmd, **kwargs):
        captured.append(cmd)
        return _Result()

    monkeypatch.setattr(restore_point.subprocess, "run", fake_run)
    restore_point.create_restore_point("Before Office cleanup")

    assert captured
    assert captured[0][0] == "powershell.exe"
