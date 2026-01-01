"""!
@brief Tests for filesystem helper utilities.
@details Validates discovery, whitelist handling, backup logic, and default
directory resolution implemented in :mod:`office_janitor.fs_tools`.
"""

from __future__ import annotations

import pathlib
import sys
import types

PROJECT_ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import fs_tools  # noqa: E402


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


def test_schedule_delete_on_reboot_uses_movefileex(tmp_path, monkeypatch) -> None:
    """!
    @brief ``MoveFileExW`` should be used when available to queue deletions.
    """

    class _FakeMoveFileEx:
        def __init__(self) -> None:
            self.calls: list[tuple[str, str | None, int]] = []
            self.argtypes = ()
            self.restype = None

        def __call__(self, source: str, destination: str | None, flags: int) -> int:
            self.calls.append((source, destination, flags))
            return 1

    fake_move = _FakeMoveFileEx()

    class _FakeKernel32:
        def __init__(self, move) -> None:
            self.MoveFileExW = move

    fake_ctypes = types.SimpleNamespace(
        windll=types.SimpleNamespace(kernel32=_FakeKernel32(fake_move)),
        c_wchar_p=str,
        c_uint=int,
        c_int=int,
        get_last_error=lambda: 0,
    )

    monkeypatch.setattr(fs_tools, "ctypes", fake_ctypes)
    monkeypatch.setattr(fs_tools.os, "name", "nt")
    monkeypatch.setattr(fs_tools.logging_ext, "get_human_logger", lambda: _NullLogger())
    monkeypatch.setattr(fs_tools.logging_ext, "get_machine_logger", lambda: _NullLogger())

    target = tmp_path / "locked.txt"
    result = fs_tools._schedule_delete_on_reboot(target)

    assert result is True
    assert fake_move.calls == [(str(target), None, fs_tools._MOVEFILE_DELAY_UNTIL_REBOOT)]


def test_schedule_delete_on_reboot_registry_fallback(tmp_path, monkeypatch) -> None:
    """!
    @brief Registry updates should be used when ``MoveFileExW`` is unavailable.
    """

    class _FakeHandle:
        def __init__(self, registry) -> None:
            self._registry = registry

        def __enter__(self):
            return self

        def __exit__(self, exc_type, exc, tb) -> bool:
            return False

    class _FakeWinReg:
        def __init__(self) -> None:
            self.values: dict[str, tuple[int, list[str]]] = {}
            self.HKEY_LOCAL_MACHINE = object()
            self.KEY_READ = 0x1
            self.KEY_SET_VALUE = 0x2
            self.REG_MULTI_SZ = 7

        def ConnectRegistry(self, _machine, _root):
            return _FakeHandle(self)

        def OpenKey(self, _registry, _subkey, _reserved, _access):
            return _FakeHandle(self)

        def QueryValueEx(self, _key, value_name):
            if value_name in self.values:
                stored = self.values[value_name]
                return stored[1], stored[0]
            raise FileNotFoundError

        def SetValueEx(self, _key, value_name, _reserved, regtype, value):
            self.values[value_name] = (regtype, list(value))

    fake_winreg = _FakeWinReg()

    monkeypatch.setattr(fs_tools, "ctypes", types.SimpleNamespace(windll=types.SimpleNamespace()))
    monkeypatch.setattr(fs_tools, "winreg", fake_winreg)
    monkeypatch.setattr(fs_tools.os, "name", "nt")
    monkeypatch.setattr(fs_tools.logging_ext, "get_human_logger", lambda: _NullLogger())
    monkeypatch.setattr(fs_tools.logging_ext, "get_machine_logger", lambda: _NullLogger())

    target = tmp_path / "delayed"
    result = fs_tools._schedule_delete_on_reboot(target)

    assert result is True
    stored = fake_winreg.values[fs_tools._PENDING_FILE_RENAME_VALUE][1]
    assert stored[-2:] == [str(target), ""]


def test_remove_paths_schedules_on_permission_error(tmp_path, monkeypatch) -> None:
    """!
    @brief Removal failures should queue a deferred deletion request.
    """

    target = tmp_path / "blocked.txt"
    target.write_text("payload", encoding="utf-8")

    calls: list[pathlib.Path] = []

    def _fake_schedule(path: pathlib.Path, *, dry_run: bool, human_logger, machine_logger) -> bool:
        calls.append(path)
        return True

    original_unlink = pathlib.Path.unlink

    def _failing_unlink(self, *args, **kwargs):  # pragma: no cover - signature passthrough
        if self == target:
            raise PermissionError("locked")
        return original_unlink(self, *args, **kwargs)

    monkeypatch.setattr(fs_tools, "make_paths_writable", lambda _paths, **_kwargs: None)
    monkeypatch.setattr(fs_tools, "reset_acl", lambda _path: None)
    monkeypatch.setattr(fs_tools, "_schedule_delete_on_reboot", _fake_schedule)
    monkeypatch.setattr(pathlib.Path, "unlink", _failing_unlink)
    monkeypatch.setattr(fs_tools.logging_ext, "get_human_logger", lambda: _NullLogger())
    monkeypatch.setattr(fs_tools.logging_ext, "get_machine_logger", lambda: _NullLogger())

    fs_tools.remove_paths([target])

    assert calls == [target]
    assert target.exists()


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
