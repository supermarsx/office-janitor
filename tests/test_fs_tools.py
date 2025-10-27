"""!
@brief Tests for filesystem helper utilities.
@details Validates discovery, whitelist handling, backup logic, and default
directory resolution implemented in :mod:`office_janitor.fs_tools`.
"""

from __future__ import annotations

import pathlib
import sys

PROJECT_ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import fs_tools


class _NullLogger:
    """!
    @brief Minimal logger stand-in discarding all messages.
    """

    def __getattr__(self, _name):  # pragma: no cover - trivial passthrough
        return lambda *args, **kwargs: None


def test_discover_paths_filters_duplicates(tmp_path) -> None:
    """!
    @brief Only existing whitelisted paths should be discovered once.
    """

    target = tmp_path / "Microsoft" / "Office"
    target.mkdir(parents=True)
    missing = tmp_path / "Missing"

    whitelist = (str(target), str(missing))
    blacklist: tuple[str, ...] = ()

    discovered = fs_tools.discover_paths(
        [target, missing, str(target)],
        whitelist=whitelist,
        blacklist=blacklist,
    )

    assert discovered == [target]


def test_filter_whitelisted_paths(tmp_path) -> None:
    """!
    @brief Ensure helper removes entries outside the whitelist.
    """

    allowed = tmp_path / "Allowed"
    blocked = tmp_path / "Blocked"
    allowed.mkdir()
    blocked.mkdir()

    filtered = fs_tools.filter_whitelisted_paths(
        [allowed, blocked],
        whitelist=(str(allowed),),
        blacklist=(str(blocked),),
    )

    assert filtered == [allowed]


def test_is_path_whitelisted_expands_environment() -> None:
    """!
    @brief ``%APPDATA%`` expansions should match the whitelist.
    """

    env = {"APPDATA": r"C:\\Users\\Alice\\AppData\\Roaming"}
    path = r"C:\\Users\\Alice\\AppData\\Roaming\\Microsoft\\Office"

    assert fs_tools.is_path_whitelisted(path, env=env)


def test_backup_path_copies_file(tmp_path, monkeypatch) -> None:
    """!
    @brief Backup helper should duplicate files into the destination root.
    """

    source = tmp_path / "source.txt"
    source.write_text("payload", encoding="utf-8")
    destination = tmp_path / "backups"

    monkeypatch.setattr(fs_tools.logging_ext, "get_human_logger", lambda: _NullLogger())
    monkeypatch.setattr(fs_tools.logging_ext, "get_machine_logger", lambda: _NullLogger())

    created = fs_tools.backup_path(source, destination)

    assert created is not None
    assert created.parent == destination
    assert created.read_text(encoding="utf-8") == "payload"


def test_backup_path_dry_run(tmp_path, monkeypatch) -> None:
    """!
    @brief Dry-run mode should avoid creating directories or files.
    """

    source = tmp_path / "dry.txt"
    source.write_text("payload", encoding="utf-8")
    destination = tmp_path / "backups"

    monkeypatch.setattr(fs_tools.logging_ext, "get_human_logger", lambda: _NullLogger())
    monkeypatch.setattr(fs_tools.logging_ext, "get_machine_logger", lambda: _NullLogger())

    created = fs_tools.backup_path(source, destination, dry_run=True)

    assert created is not None
    assert not destination.exists()


def test_get_default_log_directory_prefers_programdata(tmp_path) -> None:
    """!
    @brief Windows defaults should use ``ProgramData`` when available.
    """

    env = {"ProgramData": r"C:\\Data"}
    result = fs_tools.get_default_log_directory(env=env, platform="nt")

    assert str(result) == str(pathlib.Path(r"C:\\Data") / "OfficeJanitor" / "logs")


def test_get_default_log_directory_respects_xdg(tmp_path) -> None:
    """!
    @brief POSIX defaults should honour ``XDG_STATE_HOME`` when set.
    """

    env = {"XDG_STATE_HOME": str(tmp_path)}
    result = fs_tools.get_default_log_directory(env=env, platform="posix")

    assert result == tmp_path / "office-janitor" / "logs"


def test_get_default_backup_directory_override(tmp_path) -> None:
    """!
    @brief Explicit overrides should be returned without modification.
    """

    env = {"OFFICE_JANITOR_BACKUPDIR": str(tmp_path / "custom")}
    result = fs_tools.get_default_backup_directory(env=env, platform="posix")

    assert result == tmp_path / "custom"
