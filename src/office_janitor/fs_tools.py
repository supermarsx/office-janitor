"""!
@brief Filesystem utilities for Office residue cleanup.
@details Implements path discovery, whitelist enforcement, attribute reset,
 deletion scheduling, and backup helpers mirroring the behaviour described in
 :mod:`spec.md`. All functions rely solely on the standard library so they can
 be exercised on non-Windows hosts while still modelling the Windows-centric
 workflow used by the real scrubber.
"""

from __future__ import annotations

import ctypes
import hashlib
import os
import re
import shutil
import stat
from collections.abc import Iterable, Mapping, Sequence
from pathlib import Path
from typing import TYPE_CHECKING, Callable, cast

if TYPE_CHECKING:  # pragma: no cover - typing only
    import winreg as _winreg
else:
    try:  # pragma: no cover - platform specific module availability
        import winreg as _winreg  # type: ignore[import-not-found]
    except ImportError:  # pragma: no cover - non-Windows hosts
        class _WinRegStub:
            KEY_READ = 0
            KEY_SET_VALUE = 0
            REG_MULTI_SZ = 7
            HKEY_LOCAL_MACHINE = 0x80000002

            def ConnectRegistry(self, *args, **kwargs):  # type: ignore[no-untyped-def]
                raise FileNotFoundError

        _winreg = _WinRegStub()  # type: ignore[assignment]

winreg = _winreg

from . import constants, exec_utils, logging_ext

FILESYSTEM_WHITELIST: tuple[str, ...] = (
    r"C:\\PROGRAM FILES\\MICROSOFT OFFICE",
    r"C:\\PROGRAM FILES (X86)\\MICROSOFT OFFICE",
    r"C:\\PROGRAMDATA\\MICROSOFT\\OFFICE",
    r"C:\\PROGRAMDATA\\MICROSOFT\\CLICKTORUN",
    r"C:\\PROGRAMDATA\\MICROSOFT\\LICENSES",
    r"%LOCALAPPDATA%\\MICROSOFT\\OFFICE",
    r"%APPDATA%\\MICROSOFT\\OFFICE",
    r"%APPDATA%\\MICROSOFT\\TEMPLATES",
    r"%LOCALAPPDATA%\\MICROSOFT\\OFFICE\\TEMPLATES",
    r"%LOCALAPPDATA%\\MICROSOFT\\OFFICE\\LICENSES",
    r"%LOCALAPPDATA%\\MICROSOFT\\IDENTITYCACHE",
    r"%LOCALAPPDATA%\\MICROSOFT\\ONEAUTH",
    r"%APPDATA%\\MICROSOFT\\VBA",
    r"%LOCALAPPDATA%\\MICROSOFT\\VBA",
)
"""!
@brief Default filesystem whitelist mirroring OffScrub guardrails.
"""

FILESYSTEM_BLACKLIST: tuple[str, ...] = (
    r"C:\\WINDOWS",
    r"C:\\WINDOWS\\SYSTEM32",
    r"C:\\USERS",
)
"""!
@brief Critical system roots that must never be deleted by cleanup steps.
"""

_ENVIRONMENT_DEFAULTS: Mapping[str, str] = {
    "PROGRAMDATA": r"C:\\ProgramData",
    "LOCALAPPDATA": r"C:\\Users\\Default\\AppData\\Local",
    "APPDATA": r"C:\\Users\\Default\\AppData\\Roaming",
}
"""!
@brief Fallback values used when Windows-style variables are missing.
"""

_ENV_PATTERN = re.compile(r"%([A-Za-z0-9_]+)%")

_MOVEFILE_DELAY_UNTIL_REBOOT = 0x00000004
_PENDING_FILE_RENAME_KEY = r"SYSTEM\\CurrentControlSet\\Control\\Session Manager"
_PENDING_FILE_RENAME_VALUE = "PendingFileRenameOperations"


def normalize_windows_path(path: str | os.PathLike[str]) -> str:
    """!
    @brief Normalise a filesystem path using Windows-style casing and separators.
    @details ``Path`` objects on non-Windows hosts treat drive letters as plain
    text, so this helper focuses on canonical casing and slash collapsing rather
    than calling :func:`pathlib.Path.resolve`.
    """

    raw = str(path)
    normalized = raw.replace("/", "\\").strip()
    while "\\\\" in normalized:
        normalized = normalized.replace("\\\\", "\\")
    return normalized.upper().rstrip("\\")


def match_environment_suffix(path: str, suffix: str, *, require_users: bool = False) -> bool:
    """!
    @brief Determine whether ``path`` contains the supplied suffix segment.
    @details The helper is aware of the ``\\USERS\\`` requirement used by the
    whitelist when expanding ``%APPDATA%`` or ``%LOCALAPPDATA%`` entries.
    """

    normalized = normalize_windows_path(path)
    normalized_suffix = normalize_windows_path(suffix)
    if normalized_suffix and normalized_suffix in normalized:
        index = normalized.index(normalized_suffix)
        if require_users and "\\USERS\\" not in normalized[:index]:
            return False
        return True
    return False


def _prepare_environment(env: Mapping[str, object] | None) -> dict[str, str]:
    prepared: dict[str, str] = {}
    source = env if env is not None else os.environ
    for key, value in source.items():
        try:
            key_text = str(key).upper()
            value_text = str(value)
        except Exception:
            continue
        prepared[key_text] = value_text
    return prepared


def _lookup_env(name: str, env: Mapping[str, str]) -> str | None:
    candidate = env.get(name.upper())
    if candidate:
        return candidate
    return _ENVIRONMENT_DEFAULTS.get(name.upper())


def _expand_environment(path: str, env: Mapping[str, str]) -> str:
    def replacer(match: re.Match[str]) -> str:
        variable = match.group(1)
        value = _lookup_env(variable, env)
        return value if value is not None else match.group(0)

    return _ENV_PATTERN.sub(replacer, path)


def is_path_whitelisted(
    path: str | os.PathLike[str],
    *,
    whitelist: Iterable[str] | None = None,
    blacklist: Iterable[str] | None = None,
    env: Mapping[str, object] | None = None,
) -> bool:
    """!
    @brief Check whether ``path`` is within the allowed Office cleanup roots.
    """

    environment = _prepare_environment(env)
    allowed = tuple(whitelist or FILESYSTEM_WHITELIST)
    blocked = tuple(blacklist or FILESYSTEM_BLACKLIST)
    normalized = normalize_windows_path(path)

    for entry in allowed:
        entry_text = str(entry)
        entry_upper = entry_text.upper()
        if entry_upper.startswith("%APPDATA%\\"):
            suffix = entry_upper[len("%APPDATA%") :]
            if match_environment_suffix(
                normalized, "\\APPDATA\\ROAMING" + suffix, require_users=True
            ):
                return True
            continue
        if entry_upper.startswith("%LOCALAPPDATA%\\"):
            suffix = entry_upper[len("%LOCALAPPDATA%") :]
            if match_environment_suffix(
                normalized, "\\APPDATA\\LOCAL" + suffix, require_users=True
            ):
                return True
            continue

        expanded = _expand_environment(entry_text, environment)
        candidate = normalize_windows_path(expanded)
        if candidate and normalized.startswith(candidate):
            return True
        if normalized == candidate:
            return True

    for entry in blocked:
        expanded = _expand_environment(str(entry), environment)
        candidate = normalize_windows_path(expanded)
        if candidate and normalized.startswith(candidate):
            return False

    return False


def filter_whitelisted_paths(
    paths: Iterable[str | os.PathLike[str]],
    *,
    whitelist: Iterable[str] | None = None,
    blacklist: Iterable[str] | None = None,
    env: Mapping[str, object] | None = None,
) -> list[Path]:
    """!
    @brief Return the subset of ``paths`` permitted by the whitelist/blacklist rules.
    """

    environment = _prepare_environment(env)
    allowed_paths: list[Path] = []
    seen: set[str] = set()
    for raw in paths:
        candidate_path = Path(raw)
        normalized = normalize_windows_path(candidate_path)
        if normalized in seen:
            continue
        if is_path_whitelisted(
            candidate_path,
            whitelist=whitelist,
            blacklist=blacklist,
            env=environment,
        ):
            allowed_paths.append(candidate_path)
            seen.add(normalized)
    return allowed_paths


def discover_paths(
    candidates: Iterable[str | os.PathLike[str]] | None = None,
    *,
    whitelist: Iterable[str] | None = None,
    blacklist: Iterable[str] | None = None,
    env: Mapping[str, object] | None = None,
    must_exist: bool = True,
) -> list[Path]:
    """!
    @brief Locate filesystem entries worth cleaning based on known templates.
    @details When ``candidates`` is omitted, values from
    :data:`constants.INSTALL_ROOT_TEMPLATES` and :data:`FILESYSTEM_WHITELIST`
    seed the search.
    """

    environment = _prepare_environment(env)
    search_roots: list[str | os.PathLike[str]] = []
    if candidates is None:
        search_roots.extend(template["path"] for template in constants.INSTALL_ROOT_TEMPLATES)
        search_roots.extend(FILESYSTEM_WHITELIST)
    else:
        search_roots.extend(candidates)

    allowed = tuple(whitelist or FILESYSTEM_WHITELIST)
    blocked = tuple(blacklist or FILESYSTEM_BLACKLIST)

    discovered: list[Path] = []
    seen: set[str] = set()
    for entry in search_roots:
        expanded = _expand_environment(str(entry), environment)
        candidate_path = Path(expanded)
        normalized = normalize_windows_path(expanded)
        if normalized in seen:
            continue
        if not is_path_whitelisted(
            candidate_path,
            whitelist=allowed,
            blacklist=blocked,
            env=environment,
        ):
            continue
        exists = True
        if must_exist:
            try:
                exists = candidate_path.exists()
            except OSError:
                exists = False
        if not exists:
            continue
        discovered.append(candidate_path)
        seen.add(normalized)

    return discovered


def _handle_readonly(
    function, path: str, exc_info
) -> None:  # pragma: no cover - defensive callback
    """!
    @brief Clear read-only attributes before retrying removal.
    """

    if isinstance(exc_info[1], PermissionError):
        os.chmod(path, stat.S_IWRITE)
        function(path)
    else:
        raise exc_info[1]


def _get_movefileex() -> Callable[[str, str | None, int], int] | None:
    """!
    @brief Retrieve the Windows ``MoveFileExW`` API when available.
    @details Returns ``None`` on non-Windows hosts or when ``ctypes`` cannot
    expose the API entry point.
    """

    if os.name != "nt":
        return None
    if ctypes is None:  # pragma: no cover - defensive guard
        return None
    try:
        movefileex = ctypes.windll.kernel32.MoveFileExW  # type: ignore[attr-defined]
    except AttributeError:
        return None
    try:  # pragma: no cover - attribute assignment skipped in tests
        movefileex.argtypes = (ctypes.c_wchar_p, ctypes.c_wchar_p, ctypes.c_uint)
        movefileex.restype = ctypes.c_int
    except AttributeError:
        pass
    return movefileex


def _queue_pending_file_rename(path: str) -> bool:
    """!
    @brief Append ``path`` to ``PendingFileRenameOperations``.
    @details Acts as a fallback when :func:`MoveFileExW` is unavailable by
    updating the registry multi-string value used by Windows to process delayed
    deletions after a reboot.
    """

    if winreg is None:
        return False

    try:
        access = winreg.KEY_READ | winreg.KEY_SET_VALUE  # type: ignore[attr-defined]
    except AttributeError:  # pragma: no cover - legacy Python
        access = getattr(winreg, "KEY_WRITE", 0)

    try:
        with winreg.ConnectRegistry(None, winreg.HKEY_LOCAL_MACHINE) as registry:  # type: ignore[attr-defined]
            with winreg.OpenKey(registry, _PENDING_FILE_RENAME_KEY, 0, access) as key:
                try:
                    existing, regtype = winreg.QueryValueEx(key, _PENDING_FILE_RENAME_VALUE)
                except FileNotFoundError:
                    existing = []
                    regtype = getattr(winreg, "REG_MULTI_SZ", 7)
                if regtype != getattr(winreg, "REG_MULTI_SZ", 7):
                    existing = []
                queued = list(existing)
                queued.extend([path, ""])
                winreg.SetValueEx(
                    key,
                    _PENDING_FILE_RENAME_VALUE,
                    0,
                    getattr(winreg, "REG_MULTI_SZ", 7),
                    queued,
                )
                return True
    except OSError:
        return False
    return False


def _schedule_delete_on_reboot(
    path: Path,
    *,
    dry_run: bool = False,
    human_logger=None,
    machine_logger=None,
) -> bool:
    """!
    @brief Queue ``path`` for removal during the next system reboot.
    @details The helper first attempts to call ``MoveFileExW`` via ``ctypes``.
    When that fails it falls back to updating the ``PendingFileRenameOperations``
    registry value. A ``True`` return value indicates that one of the strategies
    accepted the request.
    """

    human_logger = human_logger or logging_ext.get_human_logger()
    machine_logger = machine_logger or logging_ext.get_machine_logger()
    path_text = str(path)

    if dry_run:
        human_logger.info("Dry-run: would schedule %s for deletion on reboot", path_text)
        machine_logger.info(
            "filesystem_remove_queued",
            extra={
                "event": "filesystem_remove_queued",
                "path": path_text,
                "method": "dry_run",
                "dry_run": True,
            },
        )
        return True

    movefileex = _get_movefileex()
    if movefileex is not None:
        try:
            result = movefileex(path_text, None, _MOVEFILE_DELAY_UNTIL_REBOOT)
        except Exception as exc:  # pragma: no cover - defensive logging
            human_logger.warning("MoveFileExW failed for %s: %s", path_text, exc)
        else:
            if result:
                human_logger.info("Queued %s for deletion on next reboot via MoveFileEx", path_text)
                machine_logger.info(
                    "filesystem_remove_queued",
                    extra={
                        "event": "filesystem_remove_queued",
                        "path": path_text,
                        "method": "movefileex",
                        "dry_run": False,
                    },
                )
                return True
            get_last_error = getattr(ctypes, "get_last_error", lambda: None)
            human_logger.warning(
                "MoveFileExW did not queue %s for deletion (error %s)",
                path_text,
                get_last_error(),
            )

    if _queue_pending_file_rename(path_text):
        human_logger.info(
            "Queued %s for deletion on next reboot via PendingFileRenameOperations",
            path_text,
        )
        machine_logger.info(
            "filesystem_remove_queued",
            extra={
                "event": "filesystem_remove_queued",
                "path": path_text,
                "method": "registry",
                "dry_run": False,
            },
        )
        return True

    human_logger.warning("Unable to queue %s for deletion on reboot", path_text)
    machine_logger.warning(
        "filesystem_remove_queue_failed",
        extra={
            "event": "filesystem_remove_queue_failed",
            "path": path_text,
            "dry_run": False,
        },
    )
    return False


def remove_paths(paths: Iterable[Path | str], *, dry_run: bool = False) -> None:
    """!
    @brief Delete the supplied paths recursively while respecting dry-run behaviour.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    for raw in paths:
        target = Path(raw)
        machine_logger.info(
            "filesystem_remove_plan",
            extra={
                "event": "filesystem_remove_plan",
                "path": str(target),
                "dry_run": bool(dry_run),
            },
        )

        if dry_run:
            human_logger.info("Dry-run: would remove %s", target)
            continue

        try:
            make_paths_writable([target])
            reset_acl(target)
        except Exception as exc:  # pragma: no cover - logged for diagnostics
            human_logger.warning("Unable to reset ACLs for %s: %s", target, exc)

        try:
            exists = target.exists()
        except OSError:
            exists = False

        if not exists:
            human_logger.debug("Skipping %s because it does not exist", target)
            continue

        human_logger.info("Removing %s", target)
        if target.is_dir():
            try:
                shutil.rmtree(target, onerror=_handle_readonly)
            except PermissionError as exc:
                human_logger.warning("Unable to remove %s due to permissions: %s", target, exc)
                _schedule_delete_on_reboot(
                    target,
                    dry_run=dry_run,
                    human_logger=human_logger,
                    machine_logger=machine_logger,
                )
            except OSError as exc:  # pragma: no cover - unexpected failure
                human_logger.warning("Unable to remove %s: %s", target, exc)
                _schedule_delete_on_reboot(
                    target,
                    dry_run=dry_run,
                    human_logger=human_logger,
                    machine_logger=machine_logger,
                )
        else:
            try:
                target.unlink()
            except PermissionError:
                os.chmod(target, stat.S_IWRITE)
                try:
                    target.unlink()
                except PermissionError as exc:
                    human_logger.warning("Unable to remove %s due to permissions: %s", target, exc)
                    _schedule_delete_on_reboot(
                        target,
                        dry_run=dry_run,
                        human_logger=human_logger,
                        machine_logger=machine_logger,
                    )
                except OSError as exc:  # pragma: no cover - unexpected failure
                    human_logger.warning("Unable to remove %s: %s", target, exc)
                    _schedule_delete_on_reboot(
                        target,
                        dry_run=dry_run,
                        human_logger=human_logger,
                        machine_logger=machine_logger,
                    )
            except OSError as exc:
                human_logger.warning("Unable to remove %s: %s", target, exc)
                _schedule_delete_on_reboot(
                    target,
                    dry_run=dry_run,
                    human_logger=human_logger,
                    machine_logger=machine_logger,
                )


def reset_acl(path: Path) -> None:
    """!
    @brief Reset permissions on ``path`` so cleanup operations can proceed.
    """

    human_logger = logging_ext.get_human_logger()

    path_text = path.as_posix()
    command = [
        "icacls",
        path_text,
        "/reset",
        "/t",
        "/c",
    ]

    result = exec_utils.run_command(
        command,
        event="fs_reset_acl",
        timeout=120,
        extra={"path": path_text},
    )

    if result.returncode == 127:
        human_logger.debug("icacls is not available; skipping ACL reset for %s", path_text)
        return

    if result.returncode != 0 or result.error:
        human_logger.warning(
            "icacls reported exit code %s for %s: %s",
            result.returncode,
            path_text,
            result.stderr.strip(),
        )


def make_paths_writable(paths: Sequence[Path | str], *, dry_run: bool = False) -> None:
    """!
    @brief Clear read-only attributes in preparation for recursive deletion.
    @details Mirrors the ``attrib -R`` behaviour from ``OfficeScrubberAIO.cmd`` to
    ensure licensing and Click-to-Run directories can be removed regardless of
    inherited ACLs.
    """

    human_logger = logging_ext.get_human_logger()

    for raw in paths:
        target = Path(raw)
        for suffix in ("", "\\*"):
            command = [
                "attrib.exe",
                "-R",
                f"{target}{suffix}",
                "/S",
                "/D",
            ]
            result = exec_utils.run_command(
                command,
                event="fs_clear_attributes",
                timeout=60,
                dry_run=dry_run,
                human_message=f"Clearing attributes for {target}{suffix}",
                extra={"path": str(target), "pattern": suffix or ""},
            )

            if result.skipped:
                continue

            if result.returncode == 127:
                human_logger.debug(
                    "attrib.exe unavailable; skipping attribute reset for %s", target
                )
                break

            if result.returncode not in {0, 1} or result.error:
                human_logger.debug(
                    "attrib exited with %s for %s: %s",
                    result.returncode,
                    target,
                    result.stderr.strip(),
                )
                break


def _sanitize_backup_name(path: Path) -> str:
    normalized = normalize_windows_path(path)
    if not normalized:
        normalized = str(path)
    sanitized = re.sub(r"[^A-Z0-9._-]+", "_", normalized).strip("_")
    if not sanitized:
        sanitized = "backup"
    digest = hashlib.sha1(normalized.encode("utf-8")).hexdigest()[:8]
    return f"{sanitized}-{digest}"


def _derive_backup_destination(root: Path, source: Path) -> Path:
    stem = _sanitize_backup_name(source)
    if source.is_file():
        destination = root / f"{stem}{source.suffix}"
    else:
        destination = root / stem

    counter = 1
    while destination.exists():
        if source.is_file():
            destination = destination.with_name(f"{destination.stem}-{counter}{destination.suffix}")
        else:
            destination = destination.with_name(f"{destination.name}-{counter}")
        counter += 1
    return destination


def backup_path(
    path: Path | str, destination_root: Path | str, *, dry_run: bool = False
) -> Path | None:
    """!
    @brief Copy ``path`` into ``destination_root`` while preserving metadata.
    """

    source = Path(path)
    root = Path(destination_root)
    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    machine_logger.info(
        "filesystem_backup_plan",
        extra={
            "event": "filesystem_backup_plan",
            "source": str(source),
            "destination": str(root),
            "dry_run": bool(dry_run),
        },
    )

    try:
        exists = source.exists()
    except OSError:
        exists = False

    if not exists:
        human_logger.debug("Skipping backup for %s; path not found", source)
        return None

    destination = _derive_backup_destination(root, source)

    if dry_run:
        human_logger.info("Dry-run: would back up %s to %s", source, destination)
        return destination

    if source.is_dir():
        human_logger.info("Backing up directory %s to %s", source, destination)
        root.mkdir(parents=True, exist_ok=True)
        shutil.copytree(source, destination)
    else:
        human_logger.info("Backing up file %s to %s", source, destination)
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, destination)

    machine_logger.info(
        "filesystem_backup_complete",
        extra={
            "event": "filesystem_backup_complete",
            "source": str(source),
            "destination": str(destination),
        },
    )

    return destination


def get_default_log_directory(
    env: Mapping[str, object] | None = None,
    *,
    platform: str | None = None,
) -> Path:
    """!
    @brief Compute the default log directory taking environment hints into account.
    """

    environment = _prepare_environment(env)
    override = environment.get("OFFICE_JANITOR_LOGDIR")
    if override:
        return Path(override).expanduser()

    system = platform or os.name
    if system == "nt":
        program_data = (
            _lookup_env("PROGRAMDATA", environment) or _ENVIRONMENT_DEFAULTS["PROGRAMDATA"]
        )
        return Path(program_data) / "OfficeJanitor" / "logs"

    xdg_state = environment.get("XDG_STATE_HOME")
    if xdg_state:
        return Path(xdg_state).expanduser() / "office-janitor" / "logs"

    return Path.cwd() / "logs"


def get_default_backup_directory(
    env: Mapping[str, object] | None = None,
    *,
    platform: str | None = None,
) -> Path:
    """!
    @brief Compute the default backup directory respecting overrides and platform.
    """

    environment = _prepare_environment(env)
    override = environment.get("OFFICE_JANITOR_BACKUPDIR")
    if override:
        return Path(override).expanduser()

    system = platform or os.name
    if system == "nt":
        program_data = (
            _lookup_env("PROGRAMDATA", environment) or _ENVIRONMENT_DEFAULTS["PROGRAMDATA"]
        )
        return Path(program_data) / "OfficeJanitor" / "backups"

    xdg_state = environment.get("XDG_STATE_HOME")
    if xdg_state:
        return Path(xdg_state).expanduser() / "office-janitor" / "backups"

    return Path.cwd() / "backups"


__all__ = [
    "FILESYSTEM_BLACKLIST",
    "FILESYSTEM_WHITELIST",
    "backup_path",
    "discover_paths",
    "filter_whitelisted_paths",
    "get_default_backup_directory",
    "get_default_log_directory",
    "is_path_whitelisted",
    "make_paths_writable",
    "match_environment_suffix",
    "normalize_windows_path",
    "remove_paths",
    "reset_acl",
]
