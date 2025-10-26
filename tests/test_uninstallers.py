"""!
@brief Command composition and dry-run tests for uninstall helpers.
@details Validates that the OffScrub wrappers constructed by
:mod:`msi_uninstall` and :mod:`c2r_uninstall` mirror the reference
``OfficeScrubber.cmd`` command lines, honour dry-run semantics, and raise when
helpers fail.
"""

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


def test_msi_uninstall_invokes_offscrub(monkeypatch, tmp_path) -> None:
    """!
    @brief Ensure MSI uninstall helper constructs the OffScrub command line.
    """

    logging_ext.setup_logging(tmp_path)
    calls: List[List[str]] = []

    def fake_run(cmd, **kwargs):
        calls.append(cmd)
        return _Result()

    monkeypatch.setattr(msi_uninstall.subprocess, "run", fake_run)
    record = {"product_code": "{91160000-0011-0000-0000-0000000FF1CE}", "version": "2019"}
    msi_uninstall.uninstall_products([record])

    assert calls, "Expected OffScrub helper to be invoked"
    command = calls[0]
    assert command[0].lower().endswith("cscript.exe")
    assert command[1] == "//NoLogo"
    assert command[2].endswith("OffScrub_O16msi.vbs")
    for arg in msi_uninstall.OFFSCRUB_BASE_ARGS:
        assert arg in command
    assert any(item.startswith("/PRODUCTCODE=") for item in command)


def test_msi_uninstall_respects_dry_run(monkeypatch, tmp_path) -> None:
    """!
    @brief Dry-run mode should skip OffScrub execution.
    """

    logging_ext.setup_logging(tmp_path)
    called = False

    def fake_run(cmd, **kwargs):
        nonlocal called
        called = True
        return _Result()

    monkeypatch.setattr(msi_uninstall.subprocess, "run", fake_run)
    msi_uninstall.uninstall_products(
        [{"product_code": "{91160000-0011-0000-0000-0000000FF1CE}", "version": "2016"}],
        dry_run=True,
    )

    assert not called, "Dry-run should not invoke OffScrub"


def test_msi_uninstall_reports_failures(monkeypatch, tmp_path) -> None:
    """!
    @brief Non-zero exit codes should raise an informative error.
    """

    logging_ext.setup_logging(tmp_path)

    def fake_run(cmd, **kwargs):
        return _Result(returncode=1603)

    monkeypatch.setattr(msi_uninstall.subprocess, "run", fake_run)

    with pytest.raises(RuntimeError) as excinfo:
        msi_uninstall.uninstall_products([{"product_code": "{BAD-CODE}", "version": "2013"}])

    assert "BAD-CODE" in str(excinfo.value)


def test_c2r_uninstall_constructs_command(monkeypatch, tmp_path) -> None:
    """!
    @brief Validate Click-to-Run OffScrub command composition and logging.
    """

    logging_ext.setup_logging(tmp_path)
    calls: List[List[str]] = []

    def fake_run(cmd, **kwargs):
        calls.append(cmd)
        return _Result()

    monkeypatch.setattr(c2r_uninstall.subprocess, "run", fake_run)

    config = {"release_ids": ["O365ProPlusRetail", "VisioProRetail"]}
    c2r_uninstall.uninstall_products(config)

    assert calls, "Expected OffScrubC2R helper to be invoked"
    command = calls[0]
    assert command[0].lower().endswith("cscript.exe")
    assert command[2].endswith("OffScrubC2R.vbs")
    assert set(c2r_uninstall.OFFSCRUB_C2R_ARGS).issubset(command)
    assert any(item.startswith("/PRODUCTS=") for item in command)


def test_c2r_uninstall_dry_run(monkeypatch, tmp_path) -> None:
    """!
    @brief Dry-run skips OffScrubC2R execution.
    """

    logging_ext.setup_logging(tmp_path)
    called = False

    def fake_run(cmd, **kwargs):
        nonlocal called
        called = True
        return _Result()

    monkeypatch.setattr(c2r_uninstall.subprocess, "run", fake_run)
    c2r_uninstall.uninstall_products({"release_ids": ["Test"]}, dry_run=True)

    assert not called, "Dry-run should not invoke OffScrubC2R"
