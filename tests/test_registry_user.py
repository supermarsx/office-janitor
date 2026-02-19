"""!
@brief Tests for user-hive registry cleanup helpers.
@details Covers backup generation, dry-run behavior, and multi-user taskband
cleanup branches implemented in :mod:`office_janitor.registry_user`.
"""

from __future__ import annotations

import pathlib
import sys
from collections.abc import Sequence

PROJECT_ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import exec_utils, registry_tools, registry_user  # noqa: E402


def _command_result(
    command: Sequence[str],
    *,
    returncode: int = 0,
    skipped: bool = False,
) -> exec_utils.CommandResult:
    """!
    @brief Build a command result stub for subprocess mocking.
    """

    return exec_utils.CommandResult(
        command=[str(part) for part in command],
        returncode=returncode,
        stdout="",
        stderr="",
        duration=0.0,
        skipped=skipped,
    )


def test_cleanup_vnext_identity_registry_dry_run_reports_backup_plan(tmp_path, monkeypatch) -> None:
    """!
    @brief Dry-run should still report planned registry backups.
    """

    monkeypatch.setattr(registry_tools, "key_exists", lambda *args, **kwargs: False)

    result = registry_user.cleanup_vnext_identity_registry(
        dry_run=True,
        backup_destination=tmp_path,
    )

    assert result["backup_requested"] is True
    assert result["backup_performed"] is False
    assert result["backup_artifacts"]
    assert str(tmp_path) in str(result["backup_destination"])


def test_cleanup_vnext_identity_registry_exports_backups(tmp_path, monkeypatch) -> None:
    """!
    @brief Live cleanup should export backups before deleting keys.
    """

    commands: list[list[str]] = []

    monkeypatch.setattr(registry_tools, "key_exists", lambda *args, **kwargs: False)
    monkeypatch.setattr(registry_user.shutil, "which", lambda exe: "reg.exe")

    def fake_run(command, *, event, dry_run=False, check=False, extra=None, **kwargs):
        commands.append([str(part) for part in command])
        return _command_result(command)

    monkeypatch.setattr(registry_user.exec_utils, "run_command", fake_run)

    result = registry_user.cleanup_vnext_identity_registry(
        dry_run=False,
        backup_destination=tmp_path,
    )

    assert commands
    assert commands[0][:2] == ["reg.exe", "export"]
    assert result["backup_requested"] is True
    assert result["backup_performed"] is True
    assert result["backup_artifacts"]


def test_cleanup_taskband_registry_include_all_users_with_backup(tmp_path, monkeypatch) -> None:
    """!
    @brief Include-all-users flow should back up and process each user SID.
    """

    deleted: list[str] = []
    unloaded: list[bool] = []

    monkeypatch.setattr(registry_user.shutil, "which", lambda exe: None)
    monkeypatch.setattr(
        registry_user,
        "delete_registry_value",
        lambda key_path, value_name, **kwargs: deleted.append(f"{key_path}\\{value_name}") or True,
    )
    monkeypatch.setattr(
        registry_user,
        "load_user_registry_hives",
        lambda dry_run=False, logger=None: ["UserA"],
    )
    monkeypatch.setattr(
        registry_user,
        "unload_user_registry_hives",
        lambda dry_run=False, logger=None: unloaded.append(True) or 1,
    )
    monkeypatch.setattr(
        registry_tools,
        "iter_subkeys",
        lambda root, path, view="native": iter(
            [
                "S-1-5-21-1111",
                "S-1-5-18",
                ".DEFAULT",
                "S-1-5-21-2222_Classes",
                "S-1-5-21-3333",
            ]
        ),
    )

    result = registry_user.cleanup_taskband_registry(
        include_all_users=True,
        dry_run=False,
        backup_destination=tmp_path,
    )

    assert result["backup_requested"] is True
    assert result["backup_performed"] is True
    assert result["backup_artifacts"]
    assert "UserA" in result["users_processed"]
    assert "S-1-5-21-1111" in result["users_processed"]
    assert "S-1-5-21-3333" in result["users_processed"]
    assert unloaded == [True]
    assert deleted


def test_registry_backup_reports_missing_destination(monkeypatch) -> None:
    """!
    @brief Cleanup should report missing backup roots when none can be resolved.
    """

    monkeypatch.setattr(registry_user.logging_ext, "get_log_directory", lambda: None)
    monkeypatch.setattr(registry_tools, "iter_subkeys", lambda *args, **kwargs: iter(()))

    result = registry_user.cleanup_taskband_registry(
        include_all_users=False,
        dry_run=True,
        backup_destination=None,
        default_logdir=None,
    )

    assert result["backup_requested"] is True
    assert result["backup_performed"] is False
    assert result["backup_errors"]
