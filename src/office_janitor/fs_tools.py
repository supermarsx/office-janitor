"""!
@brief Filesystem utilities for Office residue cleanup.
@details Implements path discovery, whitelist enforcement, attribute reset, and
backup helpers mirroring the behaviour described in :mod:`spec.md`. All
functions rely solely on the standard library so they can be exercised on
non-Windows hosts while still modelling the Windows-centric workflow used by
the real scrubber.
"""
from __future__ import annotations

import hashlib
import os
import re
import shutil
import stat
from pathlib import Path
from typing import Iterable, Mapping, Sequence

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
            if match_environment_suffix(normalized, "\\APPDATA\\ROAMING" + suffix, require_users=True):
                return True
            continue
        if entry_upper.startswith("%LOCALAPPDATA%\\"):
            suffix = entry_upper[len("%LOCALAPPDATA%") :]
            if match_environment_suffix(normalized, "\\APPDATA\\LOCAL" + suffix, require_users=True):
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


def _handle_readonly(function, path: str, exc_info) -> None:  # pragma: no cover - defensive callback
    """!
    @brief Clear read-only attributes before retrying removal.
    """

    if isinstance(exc_info[1], PermissionError):
        os.chmod(path, stat.S_IWRITE)
        function(path)
    else:
        raise exc_info[1]


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
            shutil.rmtree(target, onerror=_handle_readonly)
        else:
            try:
                target.unlink()
            except PermissionError:
                os.chmod(target, stat.S_IWRITE)
                target.unlink()


def reset_acl(path: Path) -> None:
    """!
    @brief Reset permissions on ``path`` so cleanup operations can proceed.
    """

    human_logger = logging_ext.get_human_logger()

    command = [
        "icacls",
        str(path),
        "/reset",
        "/t",
        "/c",
    ]

    result = exec_utils.run_command(
        command,
        event="fs_reset_acl",
        timeout=120,
        extra={"path": str(path)},
    )

    if result.returncode == 127:
        human_logger.debug("icacls is not available; skipping ACL reset for %s", path)
        return

    if result.returncode != 0 or result.error:
        human_logger.warning(
            "icacls reported exit code %s for %s: %s",
            result.returncode,
            path,
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
                human_logger.debug("attrib.exe unavailable; skipping attribute reset for %s", target)
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


def backup_path(path: Path | str, destination_root: Path | str, *, dry_run: bool = False) -> Path | None:
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
        program_data = _lookup_env("PROGRAMDATA", environment) or _ENVIRONMENT_DEFAULTS["PROGRAMDATA"]
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
        program_data = _lookup_env("PROGRAMDATA", environment) or _ENVIRONMENT_DEFAULTS["PROGRAMDATA"]
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

