"""!
@file scrub_cleanup.py
@brief Filesystem and registry cleanup routines for Office Janitor.

@details Contains the implementation of filesystem and registry cleanup steps,
including path normalization, registry key sorting, user template preservation,
and the actual cleanup operations. These helpers are called by the StepExecutor
during the cleanup phase of scrub operations.
"""

from __future__ import annotations

import datetime
from collections.abc import Iterable, Mapping
from pathlib import Path
from typing import Any

from . import constants, fs_tools, logging_ext, registry_tools


# ---------------------------------------------------------------------------
# Progress output helper (imported from parent)
# ---------------------------------------------------------------------------


def _scrub_progress(
    message: str,
    *,
    indent: int = 0,
    newline: bool = True,
) -> None:
    """!
    @brief Emit a progress message during scrub execution.
    @details Uses the spinner module to pause/resume so messages don't interleave
    with spinner animation. Adds optional indentation for hierarchical display.
    """
    # Import from parent to avoid circular dependency
    from . import scrub

    scrub._scrub_progress(message, indent=indent, newline=newline)


# ---------------------------------------------------------------------------
# Path normalization and sorting utilities
# ---------------------------------------------------------------------------


def normalize_string_sequence(values: object) -> list[str]:
    """!
    @brief Convert an arbitrary value into a unique, ordered list of strings.
    """

    if not isinstance(values, Iterable) or isinstance(values, (str, bytes)):
        return []

    normalised: list[str] = []
    seen: set[str] = set()
    for value in values:
        if not value:
            continue
        text = value.strip() if isinstance(value, str) else str(value).strip()
        if not text:
            continue
        normalized = fs_tools.normalize_windows_path(text)
        if normalized in seen:
            continue
        seen.add(normalized)
        normalised.append(text)
    return normalised


def sort_registry_paths_deepest_first(paths: Iterable[str]) -> list[str]:
    """!
    @brief Order registry handles so child keys are processed before parents.
    @details Ensures cleanup routines delete deeply nested keys ahead of their
    parents, mirroring OffScrub's approach and preventing ``reg delete``
    failures when a parent subtree disappears before its descendants are
    handled.
    """

    indexed = list(enumerate(paths))

    def _depth(entry: str) -> int:
        normalized = fs_tools.normalize_windows_path(entry).strip("\\")
        if not normalized:
            return 0
        return normalized.count("\\")

    indexed.sort(key=lambda item: (-_depth(item[1]), item[0]))
    return [entry for _, entry in indexed]


def normalize_option_path(value: object) -> str | None:
    """!
    @brief Convert plan metadata path entries to string form.
    """

    if isinstance(value, (str, Path)):
        return str(value)
    return None


def is_user_template_path(path: str) -> bool:
    """!
    @brief Determine whether ``path`` points at a user template directory.
    @details Mirrors :func:`safety._is_template_path` without importing private
    helpers so filesystem cleanup can independently honour preservation rules.
    """

    normalized = fs_tools.normalize_windows_path(path)
    for template in constants.USER_TEMPLATE_PATHS:
        candidate = fs_tools.normalize_windows_path(template)
        if "%" not in candidate and normalized.startswith(candidate):
            return True
        if candidate.startswith("%APPDATA%\\"):
            suffix = candidate[len("%APPDATA%") :]
            if fs_tools.match_environment_suffix(
                normalized, "\\APPDATA\\ROAMING" + suffix, require_users=True
            ):
                return True
        if candidate.startswith("%LOCALAPPDATA%\\"):
            suffix = candidate[len("%LOCALAPPDATA%") :]
            if fs_tools.match_environment_suffix(
                normalized, "\\APPDATA\\LOCAL" + suffix, require_users=True
            ):
                return True
    return False


# ---------------------------------------------------------------------------
# Filesystem cleanup
# ---------------------------------------------------------------------------


def perform_filesystem_cleanup(
    metadata: Mapping[str, object],
    context_metadata: Mapping[str, object],
    *,
    dry_run: bool,
) -> None:
    """!
    @brief Remove filesystem leftovers while preserving user templates when requested.
    @details The helper deduplicates filesystem targets, honours the ``keep_templates``
    flag propagated through the context metadata, and emits preservation messages for
    any protected template directories. Only the remaining paths are forwarded to
    :func:`fs_tools.remove_paths` so template data survives unless an explicit purge
    override is supplied. Additionally handles MSOCache, AppX package, and shortcut
    cleanup when the corresponding flags are set.
    """

    human_logger = logging_ext.get_human_logger()

    paths = normalize_string_sequence(metadata.get("paths", []))
    options = dict(context_metadata.get("options", {})) if context_metadata else {}

    # Handle extended filesystem cleanup options
    clean_msocache = bool(metadata.get("clean_msocache", False))
    clean_appx = bool(metadata.get("clean_appx", False))
    clean_shortcuts = bool(metadata.get("clean_shortcuts", False))

    # Add MSOCache paths if requested
    if clean_msocache:
        msocache_paths = [
            r"C:\MSOCache",
            str(
                Path.home()
                / "AppData"
                / "Local"
                / "Microsoft"
                / "Office"
                / "16.0"
                / "OfficeFileCache"
            ),
        ]
        for mso_path in msocache_paths:
            if mso_path not in paths:
                paths.append(mso_path)
        human_logger.info("Including MSOCache paths in cleanup")

    # Add shortcut paths if requested
    if clean_shortcuts:
        [
            str(Path.home() / "Desktop"),  # Will filter for Office .lnk files
            str(Path(r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs")),
            str(
                Path.home()
                / "AppData"
                / "Roaming"
                / "Microsoft"
                / "Windows"
                / "Start Menu"
                / "Programs"
            ),
        ]
        human_logger.info("Including Office shortcut cleanup")

    if not paths and not clean_appx and not clean_shortcuts:
        human_logger.info("No filesystem paths supplied; skipping step.")
        return

    preserve_templates = bool(
        metadata.get("preserve_templates", options.get("keep_templates", False))
    )
    purge_metadata = metadata.get("purge_templates")
    if purge_metadata is not None:
        purge_templates = bool(purge_metadata)
    else:
        purge_templates = bool(options.get("force", False) and not preserve_templates)

    preserved: list[str] = []
    cleanup_targets: list[str] = []

    for path in paths:
        if preserve_templates and not purge_templates and is_user_template_path(path):
            preserved.append(path)
            continue
        cleanup_targets.append(path)

    for template_path in preserved:
        human_logger.info("Preserving user template path %s", template_path)

    if cleanup_targets:
        _scrub_progress(f"Removing {len(cleanup_targets)} filesystem paths...", indent=3)
        fs_tools.remove_paths(cleanup_targets, dry_run=dry_run)
        _scrub_progress("Filesystem path removal complete", indent=3)
    elif paths:
        human_logger.info("All filesystem cleanup targets were preserved; nothing to remove.")

    # Handle AppX package removal if requested
    if clean_appx:
        _scrub_progress("Removing Office AppX/Store packages...", indent=3)
        try:
            removed = fs_tools.remove_office_appx_packages(dry_run=dry_run)
            _scrub_progress(f"AppX cleanup complete: {len(removed)} packages processed", indent=3)
        except Exception as exc:  # pragma: no cover - defensive
            human_logger.warning("AppX cleanup encountered an error: %s", exc)

    # Handle shortcut cleanup if requested
    if clean_shortcuts:
        _scrub_progress("Removing Office shortcuts from Start Menu and Desktop...", indent=3)
        try:
            shortcut_count = fs_tools.cleanup_office_shortcuts(dry_run=dry_run)
            _scrub_progress(
                f"Shortcut cleanup complete: {shortcut_count} shortcuts removed", indent=3
            )
        except Exception as exc:  # pragma: no cover - defensive
            human_logger.warning("Shortcut cleanup encountered an error: %s", exc)


# ---------------------------------------------------------------------------
# Registry cleanup
# ---------------------------------------------------------------------------


def perform_registry_cleanup(
    metadata: Mapping[str, object],
    *,
    dry_run: bool,
    default_backup: str | None,
    default_logdir: str | None,
) -> Mapping[str, object]:
    """!
    @brief Export and delete registry leftovers with backup awareness.
    @details Consolidates the registry cleanup logic so backup destinations are
    normalised once and deletions are skipped when no keys remain. The helper
    reuses plan metadata when provided and generates a timestamped backup path when
    only a log directory is available, mirroring the OffScrub behaviour. Returns a
    mapping describing whether a backup destination was requested or written so
    the caller can surface the information in the final summary. Additionally handles
    extended registry cleanup options like COM cleanup, typelib cleanup, etc.
    """

    human_logger = logging_ext.get_human_logger()

    keys = normalize_string_sequence(metadata.get("keys", []))
    keys = sort_registry_paths_deepest_first(keys)

    # Extract extended registry cleanup flags
    clean_wi_metadata = bool(metadata.get("clean_wi_metadata", False))
    remove_vba = bool(metadata.get("remove_vba", False))

    # Track additional cleanup operations performed
    extended_cleanups: dict[str, object] = {}

    # Handle Windows Installer metadata cleanup (Published Components)
    if clean_wi_metadata:
        _scrub_progress("Cleaning Windows Installer published components...", indent=3)
        try:
            wi_result = registry_tools.cleanup_published_components(dry_run=dry_run)
            extended_cleanups["wi_metadata"] = wi_result
            _scrub_progress("Windows Installer metadata cleanup complete", indent=3)
        except Exception as exc:  # pragma: no cover - defensive
            human_logger.warning("WI metadata cleanup error: %s", exc)
            extended_cleanups["wi_metadata_error"] = str(exc)

    # Handle VBA path cleanup
    if remove_vba:
        _scrub_progress("Cleaning VBA registry paths...", indent=3)
        try:
            vba_keys = [
                r"HKLM\Software\Microsoft\VBA",
                r"HKLM\Software\Wow6432Node\Microsoft\VBA",
            ]
            registry_tools.delete_keys(vba_keys, dry_run=dry_run)
            extended_cleanups["vba_removed"] = True
            _scrub_progress("VBA registry cleanup complete", indent=3)
        except Exception as exc:  # pragma: no cover - defensive
            human_logger.warning("VBA cleanup error: %s", exc)
            extended_cleanups["vba_error"] = str(exc)

    if not keys:
        human_logger.info("No registry keys supplied; skipping main key deletion.")
        return {
            "backup_requested": False,
            "backup_performed": False,
            "keys_processed": 0,
            **extended_cleanups,
        }

    step_backup = normalize_option_path(metadata.get("backup_destination")) or default_backup
    step_logdir = normalize_option_path(metadata.get("log_directory")) or default_logdir

    if step_backup is None and step_logdir is not None:
        timestamp = datetime.datetime.now(datetime.timezone.utc).strftime(
            "registry-backup-%Y%m%d-%H%M%S"
        )
        step_backup = str(Path(step_logdir) / timestamp)

    backup_requested = bool(step_backup)
    backup_performed = False

    if dry_run:
        human_logger.info(
            "Dry-run: would export %d registry keys to %s before deletion.",
            len(keys),
            step_backup or "(no destination)",
        )
    else:
        if step_backup is not None:
            _scrub_progress(f"Exporting {len(keys)} registry keys to backup...", indent=3)
            registry_tools.export_keys(keys, step_backup)
            backup_performed = True
            _scrub_progress(f"Registry backup complete: {step_backup}", indent=3)
        else:
            human_logger.warning("Proceeding without registry backup; no destination available.")

    _scrub_progress(f"Deleting {len(keys)} registry keys...", indent=3)
    registry_tools.delete_keys(keys, dry_run=dry_run)
    _scrub_progress("Registry key deletion complete", indent=3)

    return {
        "backup_destination": step_backup,
        "backup_requested": backup_requested,
        "backup_performed": backup_performed,
        "keys_processed": len(keys),
        **extended_cleanups,
    }
