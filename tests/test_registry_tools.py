"""!
@brief Registry tooling tests covering cleanup utilities.
@details Validates dry-run behaviour and command invocation for registry export
and delete helpers using mocked ``reg.exe`` availability.
"""

from __future__ import annotations

import pathlib
import sys
from typing import List

PROJECT_ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import registry_tools  # noqa: E402


class _Result:
    """!
    @brief Stub ``CompletedProcess`` for command tracking.
    """

    def __init__(self, returncode: int = 0) -> None:
        self.returncode = returncode


def test_delete_keys_invokes_reg_when_available(monkeypatch) -> None:
    """!
    @brief Deletion should call ``reg delete`` when the binary exists.
    """

    commands: List[List[str]] = []

    monkeypatch.setattr(registry_tools.shutil, "which", lambda exe: "reg.exe")

    def fake_run(cmd, **kwargs):
        commands.append(cmd)
        return _Result()

    monkeypatch.setattr(registry_tools.subprocess, "run", fake_run)

    registry_tools.delete_keys(["HKLM\\Software\\Contoso"], dry_run=False)

    assert commands == [["reg.exe", "delete", "HKLM\\Software\\Contoso", "/f"]]


def test_delete_keys_dry_run_skips_execution(monkeypatch) -> None:
    """!
    @brief Dry-run should avoid invoking ``reg.exe``.
    """

    monkeypatch.setattr(registry_tools.shutil, "which", lambda exe: "reg.exe")

    def fake_run(cmd, **kwargs):  # pragma: no cover - ensure not called
        raise AssertionError("reg.exe should not be invoked during dry-run")

    monkeypatch.setattr(registry_tools.subprocess, "run", fake_run)

    registry_tools.delete_keys(["HKCU\\Software\\Tailspin"], dry_run=True)


def test_export_keys_creates_placeholder_when_reg_missing(tmp_path, monkeypatch) -> None:
    """!
    @brief Exports should produce placeholder files if ``reg.exe`` is absent.
    """

    monkeypatch.setattr(registry_tools.shutil, "which", lambda exe: None)

    registry_tools.export_keys(["HKLM\\Software\\Fabrikam"], str(tmp_path))

    export_file = tmp_path / "HKLM_Software_Fabrikam.reg"
    assert export_file.exists()
    assert "Placeholder export" in export_file.read_text(encoding="utf-8")
