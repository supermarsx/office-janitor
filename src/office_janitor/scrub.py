"""!
@brief Orchestrate uninstallation, cleanup, and reporting steps.
@details The scrubber now mirrors the multi-pass behaviour of
``OfficeScrubber.cmd`` by iteratively executing MSI then Click-to-Run uninstall
steps, re-probing inventory, and continuing until no installations remain or a
pass cap is reached. Cleanup actions are deferred until the final pass so they
run once per scrub session.
"""
from __future__ import annotations

import datetime
from pathlib import Path
from typing import Iterable, Mapping, MutableMapping

from . import (
    c2r_uninstall,
    constants,
    detect,
    fs_tools,
    licensing,
    logging_ext,
    msi_uninstall,
    plan as plan_module,
    processes,
    registry_tools,
    restore_point,
    tasks_services,
)

DEFAULT_MAX_PASSES = 3
"""!
@brief Safety limit mirroring OffScrub's repeated scrub attempts.
"""

UNINSTALL_CATEGORIES = {"context", "detect", "msi-uninstall", "c2r-uninstall"}
CLEANUP_CATEGORIES = {
    "licensing-cleanup",
    "task-cleanup",
    "service-cleanup",
    "filesystem-cleanup",
    "registry-cleanup",
}


def execute_plan(
    plan: Iterable[Mapping[str, object]],
    *,
    dry_run: bool = False,
    max_passes: int | None = None,
) -> None:
    """!
    @brief Run each plan step while respecting dry-run safety requirements.
    @details The executor runs uninstall steps per pass, re-probes detection, and
    regenerates uninstall plans until either the system is clean or the maximum
    number of passes has been reached. Cleanup steps are executed once using the
    final plan, ensuring filesystem and licensing tasks do not repeat across
    passes.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    steps = [dict(step) for step in plan]
    if not steps:
        human_logger.info("No plan steps supplied; nothing to execute.")
        return

    context_step = next((step for step in steps if step.get("category") == "context"), None)
    context_metadata = dict(context_step.get("metadata", {})) if context_step else {}
    options = dict(context_metadata.get("options", {})) if context_metadata else {}

    global_dry_run = bool(dry_run or context_metadata.get("dry_run", False))
    max_pass_limit = int(
        max_passes
        or options.get("max_passes")
        or context_metadata.get("max_passes", DEFAULT_MAX_PASSES)
        or DEFAULT_MAX_PASSES
    )

    machine_logger.info(
        "scrub_plan_start",
        extra={
            "event": "scrub_plan_start",
            "step_count": len(steps),
            "dry_run": global_dry_run,
            "options": options,
            "max_passes": max_pass_limit,
        },
    )

    if global_dry_run:
        human_logger.info("Executing plan in dry-run mode; no destructive actions will occur.")
    else:
        if options.get("create_restore_point") or options.get("restore_point"):
            try:
                restore_point.create_restore_point("Office Janitor pre-cleanup")
            except Exception as exc:  # pragma: no cover - defensive logging
                human_logger.warning("Failed to create restore point: %s", exc)

        processes.terminate_office_processes(constants.DEFAULT_OFFICE_PROCESSES)
        processes.terminate_process_patterns(constants.OFFICE_PROCESS_PATTERNS)
        tasks_services.stop_services(constants.KNOWN_SERVICES)
        tasks_services.disable_tasks(constants.KNOWN_SCHEDULED_TASKS, dry_run=False)

    passes_run = 0
    base_options = dict(options)
    base_options["dry_run"] = global_dry_run

    current_plan = steps
    current_pass = int(context_metadata.get("pass_index", 1) or 1)
    final_plan = current_plan
    uninstalls_seen = _has_uninstall_steps(current_plan)

    while True:
        passes_run += 1
        machine_logger.info(
            "scrub_pass_start",
            extra={
                "event": "scrub_pass_start",
                "pass_index": current_pass,
                "dry_run": global_dry_run,
                "step_count": len(current_plan),
            },
        )

        _update_context_metadata(current_plan, current_pass, base_options, global_dry_run)
        if _has_uninstall_steps(current_plan):
            uninstalls_seen = True

        _execute_steps(current_plan, UNINSTALL_CATEGORIES, global_dry_run)

        machine_logger.info(
            "scrub_pass_complete",
            extra={
                "event": "scrub_pass_complete",
                "pass_index": current_pass,
                "dry_run": global_dry_run,
            },
        )

        if global_dry_run:
            final_plan = current_plan
            uninstalls_seen = uninstalls_seen or _has_uninstall_steps(current_plan)
            break

        if current_pass >= max_pass_limit:
            human_logger.warning(
                "Reached maximum scrub passes (%d); continuing to cleanup phase.",
                max_pass_limit,
            )
            final_plan = current_plan
            break

        inventory = detect.reprobe(base_options)
        next_plan_raw = plan_module.build_plan(inventory, base_options, pass_index=current_pass + 1)
        next_plan = [dict(step) for step in next_plan_raw]

        if not _has_uninstall_steps(next_plan):
            human_logger.info(
                "No remaining MSI or Click-to-Run installations detected after pass %d.",
                current_pass,
            )
            final_plan = next_plan
            current_pass += 1
            break

        current_plan = next_plan
        final_plan = next_plan
        current_pass += 1

    if final_plan:
        machine_logger.info(
            "scrub_cleanup_start",
            extra={
                "event": "scrub_cleanup_start",
                "pass_index": current_pass,
                "dry_run": global_dry_run,
            },
        )
        _update_context_metadata(final_plan, current_pass, base_options, global_dry_run)
        _annotate_cleanup_metadata(final_plan, base_options, uninstalls_seen)
        _execute_steps(final_plan, CLEANUP_CATEGORIES, global_dry_run)
        machine_logger.info(
            "scrub_cleanup_complete",
            extra={
                "event": "scrub_cleanup_complete",
                "pass_index": current_pass,
                "dry_run": global_dry_run,
            },
        )

    machine_logger.info(
        "scrub_plan_complete",
        extra={
            "event": "scrub_plan_complete",
            "step_count": len(final_plan) if final_plan else 0,
            "dry_run": global_dry_run,
            "passes": passes_run,
        },
    )


def _has_uninstall_steps(plan_steps: Iterable[Mapping[str, object]]) -> bool:
    """!
    @brief Determine whether a plan contains any uninstall actions.
    """

    for step in plan_steps:
        if step.get("category") in {"msi-uninstall", "c2r-uninstall"}:
            return True
    return False


def _normalize_option_path(value: object) -> str | None:
    """!
    @brief Convert plan metadata path entries to string form.
    """

    if isinstance(value, (str, Path)):
        return str(value)
    return None


def _annotate_cleanup_metadata(
    plan_steps: Iterable[MutableMapping[str, object]],
    options: Mapping[str, object],
    uninstall_detected: bool,
) -> None:
    """!
    @brief Inject runtime metadata used by cleanup steps prior to execution.
    """

    context_step: MutableMapping[str, object] | None = None
    for step in plan_steps:
        if step.get("category") == "context":
            context_step = step
            break

    context_metadata: MutableMapping[str, object] = {}
    context_options: MutableMapping[str, object] = {}
    if context_step is not None:
        context_metadata = dict(context_step.get("metadata", {}))
        context_options = dict(context_metadata.get("options", {}))

    backup_candidate = (
        context_metadata.get("backup_destination")
        or context_options.get("backup_destination")
        or context_options.get("backup")
        or options.get("backup_destination")
        or options.get("backup")
    )
    backup_path = _normalize_option_path(backup_candidate)
    force_flag = bool(options.get("force", False))

    if context_step is not None:
        context_metadata["uninstall_detected"] = uninstall_detected
        if backup_path and "backup_destination" not in context_metadata:
            context_metadata["backup_destination"] = backup_path
        if backup_path:
            context_options.setdefault("backup_destination", backup_path)
            context_options.setdefault("backup", backup_path)
        context_metadata["options"] = context_options
        context_step["metadata"] = context_metadata

    for step in plan_steps:
        if step.get("category") != "licensing-cleanup":
            continue
        metadata = dict(step.get("metadata", {}))
        metadata.setdefault("force", force_flag)
        metadata.setdefault("uninstall_detected", uninstall_detected)
        if backup_path and "backup_destination" not in metadata:
            metadata["backup_destination"] = backup_path
        if "mode" not in metadata and context_metadata.get("mode"):
            metadata["mode"] = context_metadata.get("mode")
        step["metadata"] = metadata


def _update_context_metadata(
    plan_steps: Iterable[MutableMapping[str, object]],
    pass_index: int,
    options: Mapping[str, object],
    dry_run: bool,
) -> None:
    """!
    @brief Ensure the context metadata reflects the current pass and dry-run state.
    """

    for step in plan_steps:
        if step.get("category") != "context":
            continue
        metadata = dict(step.get("metadata", {}))
        metadata["pass_index"] = int(pass_index)
        metadata["dry_run"] = bool(dry_run)
        metadata["options"] = dict(options)
        step["metadata"] = metadata
        break


def _execute_steps(
    plan_steps: Iterable[Mapping[str, object]],
    categories: Iterable[str],
    dry_run: bool,
) -> None:
    """!
    @brief Execute the subset of plan steps matching ``categories``.
    """

    def _normalize_path(value: object) -> str | None:
        if isinstance(value, (str, Path)):
            return str(value)
        return None

    selected_categories = set(categories)
    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    context_metadata: Mapping[str, object] = {}
    context_options: Mapping[str, object] = {}
    for step in plan_steps:
        if step.get("category") == "context":
            context_metadata = dict(step.get("metadata", {}))
            context_options = dict(context_metadata.get("options", {}))
            break

    backup_destination = (
        _normalize_path(context_metadata.get("backup_destination"))
        or _normalize_path(context_options.get("backup_destination"))
        or _normalize_path(context_options.get("backup"))
    )

    log_directory = (
        _normalize_path(context_metadata.get("log_directory"))
        or _normalize_path(context_options.get("log_directory"))
        or _normalize_path(context_options.get("logdir"))
    )
    if not log_directory:
        configured_logdir = logging_ext.get_log_directory()
        if configured_logdir is not None:
            log_directory = str(configured_logdir)

    for step in plan_steps:
        category = step.get("category", "unknown")
        if category not in selected_categories:
            continue
        metadata = dict(step.get("metadata", {}))
        step_dry_run = bool(dry_run or metadata.get("dry_run", False))

        machine_logger.info(
            "scrub_step_start",
            extra={
                "event": "scrub_step_start",
                "step_id": step.get("id"),
                "category": category,
                "dry_run": step_dry_run,
            },
        )

        try:
            if category == "context":
                human_logger.info("Context: %s", metadata)
            elif category == "detect":
                summary = metadata.get("summary") if isinstance(metadata, dict) else None
                if summary:
                    human_logger.info("Detection summary: %s", summary)
                else:
                    human_logger.info("Detection snapshot captured.")
            elif category == "msi-uninstall":
                product = metadata.get("product", {})
                if not product:
                    human_logger.warning("Skipping MSI uninstall step without product metadata: %s", step)
                else:
                    msi_uninstall.uninstall_products([product], dry_run=step_dry_run)
            elif category == "c2r-uninstall":
                installation = metadata.get("installation") or metadata
                if not installation:
                    human_logger.warning("Skipping C2R uninstall step without installation metadata")
                else:
                    c2r_uninstall.uninstall_products(installation, dry_run=step_dry_run)
            elif category == "licensing-cleanup":
                metadata["dry_run"] = step_dry_run
                licensing.cleanup_licenses(metadata)
            elif category == "task-cleanup":
                tasks = [str(task) for task in metadata.get("tasks", []) if task]
                if not tasks:
                    human_logger.info("No scheduled tasks supplied; skipping step.")
                else:
                    tasks_services.remove_tasks(tasks, dry_run=step_dry_run)
            elif category == "service-cleanup":
                services = [str(service) for service in metadata.get("services", []) if service]
                if not services:
                    human_logger.info("No services supplied; skipping step.")
                else:
                    tasks_services.delete_services(services, dry_run=step_dry_run)
            elif category == "filesystem-cleanup":
                _perform_filesystem_cleanup(
                    metadata,
                    context_metadata,
                    dry_run=step_dry_run,
                )
            elif category == "registry-cleanup":
                _perform_registry_cleanup(
                    metadata,
                    dry_run=step_dry_run,
                    default_backup=backup_destination,
                    default_logdir=log_directory,
                )
            else:
                human_logger.info("Unhandled plan category %s; skipping.", category)
        except Exception as exc:
            machine_logger.error(
                "scrub_step_failure",
                extra={
                    "event": "scrub_step_failure",
                    "step_id": step.get("id"),
                    "category": category,
                    "error": repr(exc),
                },
            )
            raise
        else:
            machine_logger.info(
                "scrub_step_complete",
                extra={
                    "event": "scrub_step_complete",
                    "step_id": step.get("id"),
                    "category": category,
                    "dry_run": step_dry_run,
                },
            )


def _perform_filesystem_cleanup(
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
    override is supplied.
    """

    human_logger = logging_ext.get_human_logger()

    paths = _normalize_string_sequence(metadata.get("paths", []))
    if not paths:
        human_logger.info("No filesystem paths supplied; skipping step.")
        return

    options = dict(context_metadata.get("options", {})) if context_metadata else {}
    preserve_templates = bool(
        metadata.get("preserve_templates", options.get("keep_templates", False))
    )
    purge_templates = bool(metadata.get("purge_templates", False)) or bool(
        options.get("force", False)
    )

    preserved: list[str] = []
    cleanup_targets: list[str] = []

    for path in paths:
        if preserve_templates and not purge_templates and _is_user_template_path(path):
            preserved.append(path)
            continue
        cleanup_targets.append(path)

    for template_path in preserved:
        human_logger.info("Preserving user template path %s", template_path)

    if not cleanup_targets:
        human_logger.info("All filesystem cleanup targets were preserved; nothing to remove.")
        return

    fs_tools.remove_paths(cleanup_targets, dry_run=dry_run)


def _perform_registry_cleanup(
    metadata: Mapping[str, object],
    *,
    dry_run: bool,
    default_backup: str | None,
    default_logdir: str | None,
) -> None:
    """!
    @brief Export and delete registry leftovers with backup awareness.
    @details Consolidates the registry cleanup logic so backup destinations are
    normalised once and deletions are skipped when no keys remain. The helper
    reuses plan metadata when provided and generates a timestamped backup path when
    only a log directory is available, mirroring the OffScrub behaviour.
    """

    human_logger = logging_ext.get_human_logger()

    keys = _normalize_string_sequence(metadata.get("keys", []))
    if not keys:
        human_logger.info("No registry keys supplied; skipping step.")
        return

    step_backup = _normalize_option_path(metadata.get("backup_destination")) or default_backup
    step_logdir = _normalize_option_path(metadata.get("log_directory")) or default_logdir

    if step_backup is None and step_logdir is not None:
        timestamp = datetime.datetime.now(datetime.timezone.utc).strftime("registry-backup-%Y%m%d-%H%M%S")
        step_backup = str(Path(step_logdir) / timestamp)

    if dry_run:
        human_logger.info(
            "Dry-run: would export %d registry keys to %s before deletion.",
            len(keys),
            step_backup or "(no destination)",
        )
    else:
        if step_backup is not None:
            human_logger.info(
                "Exporting %d registry keys to %s before deletion.",
                len(keys),
                step_backup,
            )
            registry_tools.export_keys(keys, step_backup)
        else:
            human_logger.warning(
                "Proceeding without registry backup; no destination available."
            )

    registry_tools.delete_keys(keys, dry_run=dry_run)


def _normalize_string_sequence(values: object) -> list[str]:
    """!
    @brief Convert an arbitrary value into a unique, ordered list of strings.
    """

    if not isinstance(values, Iterable) or isinstance(values, (str, bytes)):
        return []

    normalised: list[str] = []
    seen: set[str] = set()
    for value in values:
        text = str(value).strip()
        if not text:
            continue
        normalized = fs_tools.normalize_windows_path(text)
        if normalized in seen:
            continue
        seen.add(normalized)
        normalised.append(text)
    return normalised


def _is_user_template_path(path: str) -> bool:
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

