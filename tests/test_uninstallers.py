from __future__ import annotations

import pathlib
import sys
from typing import List

import pytest

PROJECT_ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import c2r_uninstall, logging_ext, msi_uninstall


class _Result:
    """!
    @brief Simple stand-in for :class:`subprocess.CompletedProcess`.
    """

    def __init__(self, returncode: int = 0, stdout: str = "", stderr: str = "") -> None:
        self.returncode = returncode
        self.stdout = stdout
        self.stderr = stderr


def test_msi_uninstall_invokes_msiexec(monkeypatch, tmp_path) -> None:
    """!
    @brief Ensure MSI uninstall helper constructs the correct command line.
    """

    logging_ext.setup_logging(tmp_path)
    calls: List[List[str]] = []

    def fake_run(cmd, **kwargs):
        calls.append(cmd)
        return _Result()

    monkeypatch.setattr(msi_uninstall.subprocess, "run", fake_run)
    product_code = "{91160000-0011-0000-0000-0000000FF1CE}"
    msi_uninstall.uninstall_products([product_code])

    assert calls, "Expected msiexec to be invoked"
    command = calls[0]
    assert command[0] == msi_uninstall.MSIEXEC
    assert "/x" in command
    assert product_code in command
    assert "/log" in command
    log_path = command[command.index("/log") + 1]
    assert log_path.endswith(".log")


def test_msi_uninstall_respects_dry_run(monkeypatch, tmp_path) -> None:
    """!
    @brief Dry-run mode should skip ``msiexec`` execution.
    """

    logging_ext.setup_logging(tmp_path)
    called = False

    def fake_run(cmd, **kwargs):
        nonlocal called
        called = True
        return _Result()

    monkeypatch.setattr(msi_uninstall.subprocess, "run", fake_run)
    msi_uninstall.uninstall_products(["{91160000-0011-0000-0000-0000000FF1CE}"], dry_run=True)

    assert not called, "Dry-run should not invoke msiexec"


def test_msi_uninstall_reports_failures(monkeypatch, tmp_path) -> None:
    """!
    @brief Non-zero exit codes should raise an informative error.
    """

    logging_ext.setup_logging(tmp_path)

    def fake_run(cmd, **kwargs):
        return _Result(returncode=1603)

    monkeypatch.setattr(msi_uninstall.subprocess, "run", fake_run)

    with pytest.raises(RuntimeError) as excinfo:
        msi_uninstall.uninstall_products(["{BAD-CODE}"])

    assert "BAD-CODE" in str(excinfo.value)


def test_c2r_uninstall_constructs_command(monkeypatch, tmp_path) -> None:
    """!
    @brief Validate Click-to-Run uninstall command composition and logging.
    """

    logging_ext.setup_logging(tmp_path)
    calls: List[List[str]] = []

    def fake_run(cmd, **kwargs):
        calls.append(cmd)
        return _Result()

    monkeypatch.setattr(c2r_uninstall.subprocess, "run", fake_run)

    config = {"release_ids": ["O365ProPlusRetail"]}
    c2r_uninstall.uninstall_products(config)

    assert calls, "Expected OfficeC2RClient to be invoked"
    command = calls[0]
    assert "OfficeC2RClient.exe" in command[0]
    assert any(arg.startswith("productstoremove=") for arg in command)
    assert "/log" in command


def test_c2r_uninstall_dry_run(monkeypatch, tmp_path) -> None:
    """!
    @brief Dry-run skips ``OfficeC2RClient`` execution.
    """

    logging_ext.setup_logging(tmp_path)
    called = False

    def fake_run(cmd, **kwargs):
        nonlocal called
        called = True
        return _Result()

    monkeypatch.setattr(c2r_uninstall.subprocess, "run", fake_run)
    c2r_uninstall.uninstall_products({"release_ids": ["Test"]}, dry_run=True)

    assert not called, "Dry-run should not invoke OfficeC2RClient"
