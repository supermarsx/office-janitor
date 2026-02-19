"""!
@file scrub.py
@brief Orchestrate uninstallation, cleanup, and reporting steps.

@details The scrubber now mirrors the multi-pass behaviour of
``OfficeScrubber.cmd`` by iteratively executing MSI then Click-to-Run uninstall
steps, re-probing inventory, and continuing until no installations remain or a
pass cap is reached. Cleanup actions are deferred until the final pass so they
run once per scrub session.

This module contains the main orchestration logic (execute_plan) and progress
reporting. Step execution mechanics are in scrub_executor.py and cleanup
implementations are in scrub_cleanup.py.
"""

from __future__ import annotations  # noqa: I001

import time
from collections.abc import Iterable, Mapping, MutableMapping
from pathlib import Path

from . import (
    c2r_uninstall,  # noqa: F401 - re-exported for test patching
    constants,
    detect,
    fs_tools,  # noqa: F401 - re-exported for test patching
    licensing,  # noqa: F401 - re-exported for test patching
    logging_ext,
    msi_uninstall,  # noqa: F401 - re-exported for test patching
    processes,
    registry_tools,  # noqa: F401 - re-exported for test patching
    restore_point,
    safety,
    spinner,
    tasks_services,
)
from . import (
    plan as plan_module,
)

# Re-export from scrub_cleanup for backward compatibility
from .scrub_cleanup import (  # noqa: F401
    is_user_template_path as _is_user_template_path,
    normalize_option_path as _normalize_option_path,
    normalize_string_sequence as _normalize_string_sequence,
    perform_filesystem_cleanup as _perform_filesystem_cleanup,
    perform_registry_cleanup as _perform_registry_cleanup,
    sort_registry_paths_deepest_first as _sort_registry_paths_deepest_first,
)

# Re-export from scrub_executor for backward compatibility
from .scrub_executor import (  # noqa: F401
    DEFAULT_RETRY_COUNT,
    DEFAULT_RETRY_DELAY_BASE,
    DEFAULT_RETRY_DELAY_MAX,
    FORCE_ESCALATION_ATTEMPT,
    StepExecutionError,
    StepExecutor,
    StepResult,
    _merge_reboot_details,
)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

DEFAULT_MAX_PASSES = 1
"""!
@brief Safety limit mirroring OffScrub's repeated scrub attempts.
"""

# Global start time for continuous timestamps across scrub operations
_SCRUB_START_TIME: float | None = None


# ---------------------------------------------------------------------------
# Progress output utilities
# ---------------------------------------------------------------------------


def _get_scrub_elapsed_secs() -> float:
    """!
    @brief Get elapsed seconds since scrub operation started.
    """
    global _SCRUB_START_TIME
    if _SCRUB_START_TIME is None:
        return 0.0
    return time.perf_counter() - _SCRUB_START_TIME


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
    prefix = "  " * indent
    elapsed = _get_scrub_elapsed_secs()
    timestamp = f"[{elapsed:8.2f}s] " if elapsed > 0 else ""

    spinner.pause_for_output()
    try:
        if newline:
            print(f"{timestamp}{prefix}{message}")
        else:
            print(f"{timestamp}{prefix}{message}", end="", flush=True)
    finally:
        spinner.resume_after_output()


def _scrub_ok(message: str = "") -> None:
    """!
    @brief Emit an OK status for inline progress messages.
    """
    suffix = f" {message}" if message else ""
    spinner.pause_for_output()
    try:
        print(f"[  \033[32mOK\033[0m  ]{suffix}")
    finally:
        spinner.resume_after_output()


def _scrub_fail(message: str = "") -> None:
    """!
    @brief Emit a FAILED status for inline progress messages.
    """
    suffix = f" {message}" if message else ""
    spinner.pause_for_output()
    try:
        print(f"[\033[31mFAILED\033[0m]{suffix}")
    finally:
        spinner.resume_after_output()


# ---------------------------------------------------------------------------
# Category constants
# ---------------------------------------------------------------------------

UNINSTALL_CATEGORIES = {"context", "detect", "msi-uninstall", "c2r-uninstall"}
CLEANUP_CATEGORIES = {
    "licensing-cleanup",
    "task-cleanup",
    "service-cleanup",
    "filesystem-cleanup",
    "registry-cleanup",
}


# ---------------------------------------------------------------------------
# Main execution entry point
# ---------------------------------------------------------------------------


def execute_plan(
    plan: Iterable[Mapping[str, object]],
    *,
    dry_run: bool = False,
    max_passes: int | None = None,
    start_time: float | None = None,
) -> None:
    """!
    @brief Run each plan step while respecting dry-run safety requirements.
    @details The executor runs uninstall steps per pass, re-probes detection, and
    regenerates uninstall plans until either the system is clean or the maximum
    number of passes has been reached. Cleanup steps are executed once using the
    final plan, ensuring filesystem and licensing tasks do not repeat across
    passes.
    @param start_time Optional startup timestamp for continuous timing across modules.
    """
    global _SCRUB_START_TIME
    # Use provided start_time for continuous timestamps, or start fresh
    _SCRUB_START_TIME = start_time if start_time is not None else time.perf_counter()

    # Start the spinner thread for persistent status display
    spinner.start_spinner_thread()
    spinner.set_task("Initializing scrub engine")

    _scrub_progress("=" * 50)
    _scrub_progress("Scrub Execution Engine Starting")
    _scrub_progress("=" * 50)

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    steps = [dict(step) for step in plan]
    if not steps:
        _scrub_progress("No plan steps supplied; nothing to execute.")
        human_logger.info("No plan steps supplied; nothing to execute.")
        return

    _scrub_progress(f"Loaded {len(steps)} plan steps")

    all_results: list[StepResult] = []

    context_step = next((step for step in steps if step.get("category") == "context"), None)
    context_metadata = dict(context_step.get("metadata", {})) if context_step else {}
    options = dict(context_metadata.get("options", {})) if context_metadata else {}

    global_dry_run = bool(dry_run or context_metadata.get("dry_run", False))
    if not safety.should_execute_destructive_action(
        "scrub execution",
        dry_run=global_dry_run,
        force=bool(options.get("force", False)),
    ):
        global_dry_run = True
    # Resolve max_passes carefully:
    # 0 is a valid value (skip uninstall), so check for None explicitly
    if max_passes is not None:
        max_pass_limit = int(max_passes)
    elif options.get("max_passes") is not None:
        max_pass_limit = int(options["max_passes"])
    elif context_metadata.get("max_passes") is not None:
        max_pass_limit = int(context_metadata["max_passes"])
    else:
        max_pass_limit = DEFAULT_MAX_PASSES

    _scrub_progress(f"Configuration: dry_run={global_dry_run}, max_passes={max_pass_limit}")

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

    # Restore point
    should_request_restore_point = bool(
        options.get("create_restore_point") or options.get("restore_point")
    )
    if should_request_restore_point:
        _scrub_progress("Creating system restore point...", newline=False)
        try:
            restore_point.create_restore_point("Office Janitor pre-cleanup", dry_run=global_dry_run)
            _scrub_ok()
        except Exception as exc:  # pragma: no cover - defensive logging
            _scrub_fail(str(exc))
            human_logger.warning("Failed to create restore point: %s", exc)
    else:
        _scrub_progress("Restore point creation: skipped")

    # Extract skip flags for pre-scrub operations
    skip_processes = bool(options.get("skip_processes", False))
    skip_services = bool(options.get("skip_services", False))
    skip_tasks = bool(options.get("skip_tasks", False))

    # Pre-scrub process/service cleanup
    if global_dry_run:
        _scrub_progress("DRY RUN MODE - No destructive actions will occur")
        human_logger.info("Executing plan in dry-run mode; no destructive actions will occur.")
    else:
        if skip_processes:
            _scrub_progress("Process termination: skipped (--skip-processes)")
        else:
            _scrub_progress("Terminating Office processes...", newline=False)
            processes.terminate_office_processes(constants.DEFAULT_OFFICE_PROCESSES)
            processes.terminate_process_patterns(constants.OFFICE_PROCESS_PATTERNS)
            _scrub_ok()

        if skip_services:
            _scrub_progress("Service stopping: skipped (--skip-services)")
        else:
            _scrub_progress("Stopping Office services...", newline=False)
            tasks_services.stop_services(constants.KNOWN_SERVICES)
            _scrub_ok()

        if skip_tasks:
            _scrub_progress("Task disabling: skipped (--skip-tasks)")
        else:
            _scrub_progress("Disabling scheduled tasks...", newline=False)
            tasks_services.disable_tasks(constants.KNOWN_SCHEDULED_TASKS, dry_run=False)
            _scrub_ok()

    passes_run = 0
    base_options = dict(options)
    base_options["dry_run"] = global_dry_run

    current_plan = steps
    current_pass = int(context_metadata.get("pass_index", 1) or 1)
    final_plan = current_plan
    uninstalls_seen = _has_uninstall_steps(current_plan)

    # Skip uninstall passes entirely if max_passes is 0
    if max_pass_limit <= 0:
        _scrub_progress("-" * 50)
        _scrub_progress("Skipping uninstall passes (--skip-uninstall or passes=0)")
        _scrub_progress("-" * 50)
    else:
        _scrub_progress("-" * 50)
        _scrub_progress("Beginning uninstall passes")
        _scrub_progress("-" * 50)

        while True:
            passes_run += 1
            _scrub_progress(f"=== PASS {current_pass} of {max_pass_limit} ===")
            _scrub_progress(f"Steps in this pass: {len(current_plan)}", indent=1)

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
                _scrub_progress("Uninstall steps detected in plan", indent=1)

            _scrub_progress("Executing uninstall steps...", indent=1)
            try:
                pass_results = _execute_steps(current_plan, UNINSTALL_CATEGORIES, global_dry_run)
            except StepExecutionError as exc:
                _scrub_progress(f"Pass {current_pass} FAILED", indent=1)
                all_results.extend(exc.partial_results)
                _log_summary(all_results, passes_run, global_dry_run)
                raise
            else:
                all_results.extend(pass_results)

            pass_successes = sum(1 for item in pass_results if item.status == "success")
            pass_failures = sum(1 for item in pass_results if item.status == "failed")
            pass_skipped = len(pass_results) - pass_successes - pass_failures
            pass_duration = sum(
                (item.completed_at - item.started_at)
                for item in pass_results
                if item.started_at is not None
                and item.completed_at is not None
                and item.completed_at >= item.started_at
            )

            _scrub_progress(
                f"Pass {current_pass} complete: {pass_successes} success, "
                f"{pass_failures} failed, {pass_skipped} skipped ({pass_duration:.2f}s)",
                indent=1,
            )

            machine_logger.info(
                "scrub_pass_complete",
                extra={
                    "event": "scrub_pass_complete",
                    "pass_index": current_pass,
                    "dry_run": global_dry_run,
                    "successes": pass_successes,
                    "failures": pass_failures,
                    "skipped": pass_skipped,
                    "duration": round(pass_duration, 6),
                },
            )

            if global_dry_run:
                _scrub_progress("Dry run - skipping additional passes", indent=1)
                final_plan = current_plan
                uninstalls_seen = uninstalls_seen or _has_uninstall_steps(current_plan)
                break

            if current_pass >= max_pass_limit:
                _scrub_progress(f"Reached maximum passes ({max_pass_limit})", indent=1)
                _scrub_progress(
                    "ALERT: Uninstall pass limit reached; cleanup will continue with possible "
                    "leftovers still present.",
                    indent=1,
                )
                human_logger.warning(
                    "ALERT: Reached maximum scrub passes (%d); continuing to cleanup phase "
                    "with potential leftovers.",
                    max_pass_limit,
                )
                machine_logger.warning(
                    "scrub_pass_limit_reached",
                    extra={
                        "event": "scrub_pass_limit_reached",
                        "pass_index": current_pass,
                        "max_passes": max_pass_limit,
                        "dry_run": global_dry_run,
                        "cleanup_continues": True,
                    },
                )
                final_plan = current_plan
                break

            _scrub_progress("Re-probing inventory for next pass...", indent=1)
            inventory = detect.reprobe(base_options)
            next_plan_raw = plan_module.build_plan(
                inventory, base_options, pass_index=current_pass + 1
            )
            next_plan = [dict(step) for step in next_plan_raw]

            if not _has_uninstall_steps(next_plan):
                _scrub_progress(
                    "No remaining installations detected - uninstall phase complete", indent=1
                )
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
            _scrub_progress(f"Moving to pass {current_pass}...", indent=1)

    _scrub_progress("-" * 50)
    _scrub_progress("Beginning cleanup phase")
    _scrub_progress("-" * 50)

    if final_plan:
        cleanup_count = sum(1 for s in final_plan if s.get("category") in CLEANUP_CATEGORIES)
        _scrub_progress(f"Cleanup steps to process: {cleanup_count}")

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

        _scrub_progress("Executing cleanup steps...")
        # Cleanup steps should continue on failure - don't stop the whole process
        cleanup_results = _execute_steps(
            final_plan, CLEANUP_CATEGORIES, global_dry_run, continue_on_failure=True
        )
        all_results.extend(cleanup_results)

        cleanup_successes = sum(1 for r in cleanup_results if r.status == "success")
        cleanup_failures = sum(1 for r in cleanup_results if r.status == "failed")
        cleanup_duration = sum(
            (item.completed_at - item.started_at)
            for item in cleanup_results
            if item.started_at is not None
            and item.completed_at is not None
            and item.completed_at >= item.started_at
        )

        _scrub_progress(
            f"Cleanup complete: {cleanup_successes} success, "
            f"{cleanup_failures} failed ({cleanup_duration:.2f}s)"
        )

        machine_logger.info(
            "scrub_cleanup_complete",
            extra={
                "event": "scrub_cleanup_complete",
                "pass_index": current_pass,
                "dry_run": global_dry_run,
                "steps_processed": len(cleanup_results),
                "duration": round(cleanup_duration, 6),
            },
        )
    else:
        _scrub_progress("No cleanup steps to execute")

    _log_summary(all_results, passes_run, global_dry_run)

    total_successes = sum(1 for r in all_results if r.status == "success")
    total_failures = sum(1 for r in all_results if r.status == "failed")
    total_time = _get_scrub_elapsed_secs()

    _scrub_progress("=" * 50)
    _scrub_progress("SCRUB EXECUTION COMPLETE")
    _scrub_progress(f"Total steps: {len(all_results)}")
    _scrub_progress(f"Successes: {total_successes}")
    _scrub_progress(f"Failures: {total_failures}")
    _scrub_progress(f"Passes run: {passes_run}")
    _scrub_progress(f"Total time: {total_time:.2f}s")
    _scrub_progress("=" * 50)

    machine_logger.info(
        "scrub_plan_complete",
        extra={
            "event": "scrub_plan_complete",
            "step_count": len(final_plan) if final_plan else 0,
            "dry_run": global_dry_run,
            "passes": passes_run,
        },
    )


# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------


def _has_uninstall_steps(plan_steps: Iterable[Mapping[str, object]]) -> bool:
    """!
    @brief Determine whether a plan contains any uninstall actions.
    """

    for step in plan_steps:
        if step.get("category") in {"msi-uninstall", "c2r-uninstall"}:
            return True
    return False


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
    *,
    continue_on_failure: bool = False,
) -> list[StepResult]:
    """!
    @brief Execute the subset of plan steps matching ``categories``.
    @param continue_on_failure If True, continue to next step after failure instead of raising.
    @return Ordered list of :class:`StepResult` entries describing each step.
    """

    def _normalize_path(value: object) -> str | None:
        if isinstance(value, (str, Path)):
            return str(value)
        return None

    selected_categories = set(categories)

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

    selected_steps = [
        step for step in plan_steps if step.get("category", "unknown") in selected_categories
    ]

    executor = StepExecutor(
        dry_run=dry_run,
        context_metadata=context_metadata,
        backup_destination=backup_destination,
        log_directory=log_directory,
        total_steps=len(selected_steps) or 1,
    )

    results: list[StepResult] = []

    for index, step in enumerate(selected_steps, start=1):
        result = executor.run_step(step, index=index)
        results.append(result)
        if result.status == "failed":
            # Non-recoverable errors or continue_on_failure mode: log and continue
            # Recoverable errors (and not continue_on_failure): stop and allow pass retry
            if result.non_recoverable or continue_on_failure:
                logging_ext.get_human_logger().warning(
                    "Step %s failed%s, continuing to next step...",
                    result.step_id or result.category,
                    " (non-recoverable)" if result.non_recoverable else "",
                )
                continue
            raise StepExecutionError(result, results) from result.exception

    return results


def _log_summary(results: Iterable[StepResult], passes: int, dry_run: bool) -> None:
    """!
    @brief Emit a consolidated summary for the executed plan steps.
    @details Counts successful, failed, and skipped steps and highlights registry
    backup activity so operators can confirm safeguards executed as intended.
    The structured log entry mirrors the human-readable message to keep both
    channels aligned.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    result_list = list(results)
    total = len(result_list)
    successes = sum(1 for item in result_list if item.status == "success")
    failures = sum(1 for item in result_list if item.status == "failed")
    skipped = total - successes - failures
    backups_requested = sum(1 for item in result_list if bool(item.details.get("backup_requested")))
    backups_performed = sum(1 for item in result_list if bool(item.details.get("backup_performed")))

    reboot_recommended = False
    reboot_services: list[str] = []
    for item in result_list:
        if not bool(item.details.get("reboot_recommended")):
            continue
        reboot_recommended = True
        services = item.details.get("reboot_services")
        if isinstance(services, Iterable) and not isinstance(services, (str, bytes)):
            for service in services:
                text = str(service).strip()
                if text and text not in reboot_services:
                    reboot_services.append(text)

    reboot_services_sorted = reboot_services

    durations = [
        item.completed_at - item.started_at
        for item in result_list
        if item.started_at is not None
        and item.completed_at is not None
        and item.completed_at >= item.started_at
    ]
    total_duration = sum(durations) if durations else 0.0
    average_duration = (total_duration / len(durations)) if durations else 0.0

    machine_logger.info(
        "scrub_summary",
        extra={
            "event": "scrub_summary",
            "total_steps": total,
            "successes": successes,
            "failures": failures,
            "skipped": skipped,
            "passes": passes,
            "dry_run": dry_run,
            "backups_requested": backups_requested,
            "backups_performed": backups_performed,
            "total_duration": round(total_duration, 6),
            "average_duration": round(average_duration, 6),
            "reboot_recommended": reboot_recommended,
            "reboot_services": reboot_services_sorted or None,
        },
    )

    if total == 0:
        human_logger.info(
            "Scrub summary: no matching steps executed; passes=%d; dry_run=%s",
            passes,
            dry_run,
        )
    else:
        reboot_text = ""
        if reboot_recommended:
            display = ", ".join(reboot_services_sorted) or "system reboot required"
            reboot_text = f"; reboot recommended to finish stopping services: {display}"

        human_logger.info(
            (
                "Scrub summary: %d step(s) processed (%d succeeded, %d failed, %d skipped); "
                "passes=%d; dry_run=%s; registry backups=%d/%d; duration=%.3fs%s"
            ),
            total,
            successes,
            failures,
            skipped,
            passes,
            dry_run,
            backups_performed,
            backups_requested,
            total_duration,
            reboot_text,
        )
