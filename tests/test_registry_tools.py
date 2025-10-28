"""!
@brief Registry tooling tests covering cleanup utilities.
@details Validates dry-run behaviour, guardrails, and command invocation for
registry export and delete helpers using mocked ``reg.exe`` availability. The
tests also cover Office uninstall heuristics exposed by the module.
"""

from __future__ import annotations

import pathlib
import sys
from typing import Iterable, List, Tuple

import pytest

PROJECT_ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import exec_utils, registry_tools  # noqa: E402


def _command_result(command: Iterable[str], returncode: int = 0, *, skipped: bool = False) -> exec_utils.CommandResult:
    """!
    @brief Fabricate :class:`CommandResult` objects for command interception.
    """

    return exec_utils.CommandResult(
        command=[str(part) for part in command],
        returncode=returncode,
        stdout="",
        stderr="",
        duration=0.0,
        skipped=skipped,
    )


class _Recorder:
    """!
    @brief Minimal logger stub capturing emitted events.
    """

    def __init__(self) -> None:
        self.messages: List[Tuple[str, dict]] = []

    def info(self, message: str, *args, **kwargs) -> None:  # noqa: D401 - logging compatibility
        payload = kwargs.copy()
        self.messages.append((message, payload))

    def warning(self, message: str, *args, **kwargs) -> None:  # noqa: D401 - logging compatibility
        payload = kwargs.copy()
        self.messages.append((message, payload))


def test_delete_keys_invokes_reg_when_available(monkeypatch) -> None:
    """!
    @brief Deletion should call ``reg delete`` when the binary exists.
    """

    commands: List[List[str]] = []

    monkeypatch.setattr(registry_tools.shutil, "which", lambda exe: "reg.exe")

    def fake_run(command, *, event, dry_run=False, **kwargs):
        commands.append([str(part) for part in command])
        return _command_result(command, skipped=dry_run)

    monkeypatch.setattr(registry_tools.exec_utils, "run_command", fake_run)

    registry_tools.delete_keys(
        ["HKLM\\Software\\Microsoft\\Office\\Contoso"],
        dry_run=False,
        logger=_Recorder(),
    )

    assert commands == [["reg.exe", "delete", "HKLM\\SOFTWARE\\MICROSOFT\\OFFICE\\CONTOSO", "/f"]]


def test_delete_keys_dry_run_skips_execution(monkeypatch) -> None:
    """!
    @brief Dry-run should avoid invoking ``reg.exe``.
    """

    monkeypatch.setattr(registry_tools.shutil, "which", lambda exe: "reg.exe")

    calls: List[bool] = []

    def fake_run(command, *, event, dry_run=False, **kwargs):
        calls.append(dry_run)
        return _command_result(command, skipped=dry_run)

    monkeypatch.setattr(registry_tools.exec_utils, "run_command", fake_run)

    registry_tools.delete_keys(
        ["HKCU\\Software\\Microsoft\\Office\\Tailspin"],
        dry_run=True,
        logger=_Recorder(),
    )

    assert calls and all(calls)


def test_delete_keys_rejects_disallowed_paths(monkeypatch) -> None:
    """!
    @brief Guardrails should reject deletions outside the whitelist.
    """

    monkeypatch.setattr(registry_tools.shutil, "which", lambda exe: "reg.exe")

    with pytest.raises(registry_tools.RegistryError):
        registry_tools.delete_keys(["HKLM\\Software\\Contoso"], dry_run=False, logger=_Recorder())


def test_export_keys_creates_placeholder_when_reg_missing(tmp_path, monkeypatch) -> None:
    """!
    @brief Exports should produce placeholder files if ``reg.exe`` is absent.
    """

    monkeypatch.setattr(registry_tools.shutil, "which", lambda exe: None)

    recorder = _Recorder()
    exported = registry_tools.export_keys(
        ["HKLM\\Software\\Microsoft\\Office\\Fabrikam"],
        tmp_path,
        logger=recorder,
    )

    assert exported, "Expected an export path to be returned"
    export_file = exported[0]
    assert export_file.exists()
    assert "Placeholder export" in export_file.read_text(encoding="utf-8")
    assert recorder.messages and recorder.messages[0][1]["extra"]["action"] == "registry-export"


def test_export_keys_dry_run_records_intent(tmp_path, monkeypatch) -> None:
    """!
    @brief Dry-run exports should log intent without creating files.
    """

    monkeypatch.setattr(registry_tools.shutil, "which", lambda exe: "reg.exe")

    recorder = _Recorder()

    def fake_run(command, *, event, dry_run=False, **kwargs):
        return _command_result(command, skipped=dry_run)

    monkeypatch.setattr(registry_tools.exec_utils, "run_command", fake_run)

    exported = registry_tools.export_keys(
        ["HKLM\\Software\\Microsoft\\Office\\Diagnostics"],
        tmp_path,
        dry_run=True,
        logger=recorder,
    )

    assert exported
    assert not exported[0].exists()
    assert recorder.messages[0][1]["extra"]["dry_run"] is True


def test_looks_like_office_entry_matches_keywords() -> None:
    """!
    @brief The Office heuristic should recognise branded display names.
    """

    entry = {"DisplayName": "Microsoft Office 365 ProPlus", "Publisher": "Microsoft Corporation"}
    assert registry_tools.looks_like_office_entry(entry)

    unrelated = {"DisplayName": "Contoso Widget", "Publisher": "Contoso"}
    assert not registry_tools.looks_like_office_entry(unrelated)


def test_iter_office_uninstall_entries_filters_non_office(monkeypatch) -> None:
    """!
    @brief Only Office-like entries should be returned from uninstall enumeration.
    """

    roots: Iterable[Tuple[int, str]] = [(0x80000002, r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall")]

    def fake_iter_subkeys(root: int, path: str, *, view: str | None = None):
        yield from ("{90160000-0011-0000-0000-0000000FF1CE}", "ContosoApp")

    def fake_read_values(root: int, path: str, *, view: str | None = None):
        if path.endswith("ContosoApp"):
            return {"DisplayName": "Contoso", "Publisher": "Contoso"}
        return {
            "DisplayName": "Microsoft Office Professional Plus 2016",
            "Publisher": "Microsoft Corporation",
            "ProductCode": "{90160000-0011-0000-0000-0000000FF1CE}",
        }

    monkeypatch.setattr(registry_tools, "iter_subkeys", fake_iter_subkeys)
    monkeypatch.setattr(registry_tools, "read_values", fake_read_values)

    results = list(registry_tools.iter_office_uninstall_entries(roots))
    assert len(results) == 1
    hive, path, values = results[0]
    assert hive == 0x80000002
    assert path.endswith("{90160000-0011-0000-0000-0000000FF1CE}")
    assert values["ProductCode"] == "{90160000-0011-0000-0000-0000000FF1CE}"

