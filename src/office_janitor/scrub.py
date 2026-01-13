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
import time
from collections.abc import Iterable, Mapping, MutableMapping
from dataclasses import dataclass, field
from pathlib import Path

from . import (
    c2r_uninstall,
    constants,
    detect,
    fs_tools,
    licensing,
    logging_ext,
    msi_uninstall,
    processes,
    registry_tools,
    restore_point,
    tasks_services,
)
from . import (
    plan as plan_module,
)

DEFAULT_MAX_PASSES = 3
"""!
@brief Safety limit mirroring OffScrub's repeated scrub attempts.
"""

DEFAULT_RETRY_COUNT = 9
"""!
@brief Default retry count for plan steps (9 retries = 10 total attempts).
"""

DEFAULT_RETRY_DELAY_BASE = 3
"""!
@brief Base delay between retries in seconds. Progressive backoff is applied.
"""

DEFAULT_RETRY_DELAY_MAX = 30
"""!
@brief Maximum delay between retries in seconds.
"""

FORCE_ESCALATION_ATTEMPT = 3
"""!
@brief Attempt number at which force mode is enabled for uninstall operations.
"""


# ---------------------------------------------------------------------------
# Progress logging utilities
# ---------------------------------------------------------------------------

_SCRUB_START_TIME: float | None = None


def _get_scrub_elapsed_secs() -> float:
    """Return seconds since scrub started, or 0 if not started."""
    if _SCRUB_START_TIME is None:
        return 0.0
    return time.perf_counter() - _SCRUB_START_TIME


def _scrub_progress(message: str, *, newline: bool = True, indent: int = 0) -> None:
    """Print a scrub progress message with dmesg-style timestamp."""
    timestamp = f"[{_get_scrub_elapsed_secs():12.6f}]"
    prefix = "  " * indent
    if newline:
        print(f"{timestamp} {prefix}{message}", flush=True)
    else:
        print(f"{timestamp} {prefix}{message}", end="", flush=True)


def _scrub_ok(extra: str = "") -> None:
    """Print OK status in Linux init style [  OK  ]."""
    suffix = f" {extra}" if extra else ""
    print(f" [  \033[32mOK\033[0m  ]{suffix}", flush=True)


def _scrub_fail(reason: str = "") -> None:
    """Print FAIL status in Linux init style [FAILED]."""
    suffix = f" ({reason})" if reason else ""
    print(f" [\033[31mFAILED\033[0m]{suffix}", flush=True)


@dataclass
class StepResult:
    """!
    @brief Capture the outcome for a single plan step.
    @details ``status`` is one of ``"success"``, ``"failed"``, or ``"skipped"`` and is
    recorded together with the executed dry-run flag and optional detail fields
    describing retries, error messages, backup/export activity, and timeline
    metadata. The structure feeds the summary reporter at the end of a plan run
    and is also attached to retry exceptions for diagnostic logging; the
    ``exception`` field retains the final raised object for traceback chaining,
    while ``started_at``/``completed_at`` capture timing data used by the
    progress reporter.
    """

    step_id: str | None
    category: str
    status: str
    attempts: int
    dry_run: bool
    error: str | None = None
    exception: BaseException | None = None
    non_recoverable: bool = False
    progress: float | None = None
    started_at: float | None = None
    completed_at: float | None = None
    details: MutableMapping[str, object] = field(default_factory=dict)


class StepExecutionError(RuntimeError):
    """!
    @brief Raised when a plan step exhausts retries without succeeding.
    @details Carries the :class:`StepResult` describing the failure as well as
    any prior results accumulated during the batch. Callers append the partial
    results to the global execution state before propagating the exception so
    the summary reporter can emit a best-effort recap.
    """

    def __init__(self, result: StepResult, partial_results: Iterable[StepResult]):
        self.result: StepResult = result
        self.partial_results: list[StepResult] = list(partial_results)
        message = (
            f"Plan step {result.step_id or result.category} failed after "
            f"{result.attempts} attempt(s)"
        )
        super().__init__(message)


class StepExecutor:
    """!
    @brief Execute individual plan steps with retry handling and logging.
    @details The executor normalises per-step dry-run flags, emits human and
    machine readable progress updates, and performs retries when metadata
    specifies ``retries``/``retry_delay`` entries. Results are emitted as
    :class:`StepResult` instances so the caller can maintain a unified summary.
    """

    def __init__(
        self,
        *,
        dry_run: bool,
        context_metadata: Mapping[str, object],
        backup_destination: str | None,
        log_directory: str | None,
        total_steps: int,
    ) -> None:
        self._dry_run = dry_run
        self._context_metadata = context_metadata
        self._backup_destination = backup_destination
        self._log_directory = log_directory
        self._total_steps = max(1, total_steps)
        self._human_logger = logging_ext.get_human_logger()
        self._machine_logger = logging_ext.get_machine_logger()

    def run_step(self, step: Mapping[str, object], *, index: int) -> StepResult:
        """!
        @brief Execute ``step`` and return a :class:`StepResult`.
        @details Retries are sourced from ``step['retries']`` or
        ``step['metadata']['retries']``. Optional delay values are honoured via
        ``retry_delay``/``retry_delay_seconds`` keys. When an exception is
        raised, it is logged and retried until attempts are exhausted. The final
        result includes a ``details`` payload indicating whether registry
        backups occurred.
        """

        category = step.get("category", "unknown")
        step_id = step.get("id")
        metadata = dict(step.get("metadata", {}))
        dry_run = bool(self._dry_run or metadata.get("dry_run", False))
        retries = self._resolve_retry_count(step, metadata)
        delay = self._resolve_retry_delay(step, metadata)
        attempts_allowed = retries + 1

        progress = min(1.0, max(0.0, index / self._total_steps))

        result = StepResult(
            step_id=str(step_id) if step_id is not None else None,
            category=str(category),
            status="skipped",
            attempts=0,
            dry_run=dry_run,
        )
        result.progress = progress

        for attempt in range(1, attempts_allowed + 1):
            result.attempts = attempt
            start_time = time.perf_counter()
            if result.started_at is None:
                result.started_at = start_time
            self._emit_start(
                step_id,
                category,
                dry_run,
                attempt,
                index,
                progress,
            )

            try:
                detail_payload = self._dispatch(
                    category=category,
                    metadata=metadata,
                    dry_run=dry_run,
                )
            except Exception as exc:  # pragma: no cover - exercised in failure paths
                result.status = "failed"
                result.error = repr(exc)
                result.exception = exc
                result.completed_at = time.perf_counter()
                pending_reboots = tasks_services.consume_reboot_recommendations()
                if pending_reboots:
                    _merge_reboot_details(result.details, pending_reboots)

                # Extract a concise error reason for display
                error_reason = self._extract_error_reason(exc)

                # Check if this error is non-recoverable (no point retrying)
                is_non_recoverable = self._is_non_recoverable_error(exc)

                self._machine_logger.error(
                    "scrub_step_failure",
                    extra={
                        "event": "scrub_step_failure",
                        "step_id": step_id,
                        "category": category,
                        "attempt": attempt,
                        "error": repr(exc),
                        "error_reason": error_reason,
                        "non_recoverable": is_non_recoverable,
                        "progress": progress,
                        "duration": self._format_duration(result),
                    },
                )
                self._human_logger.error(
                    "Plan step %s (%s) failed on attempt %d: %s",
                    step_id or category,
                    category,
                    attempt,
                    exc,
                )

                # Skip retries for non-recoverable errors
                if is_non_recoverable:
                    result.non_recoverable = True
                    _scrub_fail(f"{error_reason} (non-recoverable, continuing)")
                    break

                if attempt < attempts_allowed:
                    # Print RETRY status with error reason
                    _scrub_fail(f"retry: {error_reason}")
                    # Calculate progressive delay with backoff
                    actual_delay = self._calculate_progressive_delay(delay, attempt)
                    # Enable force mode after FORCE_ESCALATION_ATTEMPT
                    if attempt >= FORCE_ESCALATION_ATTEMPT:
                        metadata["force"] = True
                        _scrub_progress(
                            f"Escalating to force mode after {attempt} attempts",
                            indent=3,
                        )
                    _scrub_progress(
                        f"Waiting {actual_delay}s before retry {attempt + 1}/{attempts_allowed}...",
                        indent=3,
                    )
                    time.sleep(actual_delay)
                    continue
                # Final failure - print FAILED status with error reason
                _scrub_fail(error_reason)
                break
            else:
                # Print OK status with duration
                duration = self._format_duration(result)
                duration_str = f"{duration:.2f}s" if duration is not None else ""
                _scrub_ok(duration_str)
                result.status = "success"
                result.error = None
                result.exception = None
                result.completed_at = time.perf_counter()
                if detail_payload:
                    result.details.update(detail_payload)
                pending_reboots = tasks_services.consume_reboot_recommendations()
                if pending_reboots:
                    _merge_reboot_details(result.details, pending_reboots)
                self._machine_logger.info(
                    "scrub_step_complete",
                    extra={
                        "event": "scrub_step_complete",
                        "step_id": step_id,
                        "category": category,
                        "dry_run": dry_run,
                        "attempts": attempt,
                        "progress": progress,
                        "duration": duration,
                    },
                )
                break

        return result

    def _emit_start(
        self,
        step_id: object,
        category: object,
        dry_run: bool,
        attempt: int,
        index: int,
        progress: float,
    ) -> None:
        # Print visible substep progress
        step_label = f"{step_id}" if step_id else f"{category}"
        attempt_info = f" (attempt {attempt})" if attempt > 1 else ""
        dry_run_marker = " [DRY-RUN]" if dry_run else ""
        _scrub_progress(
            f"[{index}/{self._total_steps}] {step_label}{attempt_info}{dry_run_marker}...",
            newline=False,
            indent=2,
        )
        self._human_logger.info(
            "Executing step %s/%s (%s:%s) attempt %d",
            index,
            self._total_steps,
            category,
            step_id or "<unknown>",
            attempt,
        )
        self._machine_logger.info(
            "scrub_step_start",
            extra={
                "event": "scrub_step_start",
                "step_id": step_id,
                "category": category,
                "dry_run": dry_run,
                "attempt": attempt,
                "index": index,
                "total_steps": self._total_steps,
                "progress": progress,
            },
        )

    def _format_duration(self, result: StepResult) -> float | None:
        if result.started_at is None or result.completed_at is None:
            return None
        duration = result.completed_at - result.started_at
        if duration < 0:
            return None
        return round(duration, 6)

    @staticmethod
    def _extract_error_reason(exc: Exception) -> str:
        """!
        @brief Extract a concise, human-readable error reason from an exception.
        @details Parses common exception types to provide actionable error messages.
        """
        exc_type = type(exc).__name__
        exc_msg = str(exc).strip()

        # Handle specific exception types with better messages
        if isinstance(exc, FileNotFoundError):
            return f"file not found: {exc_msg}" if exc_msg else "file not found"
        if isinstance(exc, PermissionError):
            return f"permission denied: {exc_msg}" if exc_msg else "permission denied"
        if isinstance(exc, TimeoutError):
            return "operation timed out"
        if isinstance(exc, OSError):
            # OSError can have errno
            errno_info = f" (errno {exc.errno})" if exc.errno else ""
            return f"OS error{errno_info}: {exc_msg}" if exc_msg else f"OS error{errno_info}"
        if isinstance(exc, RuntimeError):
            # RuntimeError often contains the actual reason
            return exc_msg if exc_msg else "runtime error"
        if isinstance(exc, ValueError):
            return f"invalid value: {exc_msg}" if exc_msg else "invalid value"

        # Check for subprocess/command failures
        if "return" in exc_msg.lower() and "code" in exc_msg.lower():
            return exc_msg
        if "exit" in exc_msg.lower() and "code" in exc_msg.lower():
            return exc_msg

        # Check for verification failures
        if "verification" in exc_msg.lower() or "residue" in exc_msg.lower():
            return exc_msg

        # Default: use exception message or type
        if exc_msg:
            # Truncate very long messages
            if len(exc_msg) > 80:
                return f"{exc_msg[:77]}..."
            return exc_msg
        return exc_type

    @staticmethod
    def _is_non_recoverable_error(exc: Exception) -> bool:
        """!
        @brief Determine if an exception is non-recoverable (should not be retried).
        @details Certain errors cannot be fixed by retrying - missing executables,
        invalid configuration, etc. These should fail immediately to save time.
        """
        # FileNotFoundError - missing executables, files, etc.
        if isinstance(exc, FileNotFoundError):
            return True

        # ValueError - invalid configuration, bad parameters
        if isinstance(exc, ValueError):
            return True

        # NotImplementedError - feature not available
        if isinstance(exc, NotImplementedError):
            return True

        # ImportError - missing dependencies
        if isinstance(exc, ImportError):
            return True

        # Check message for non-recoverable patterns
        exc_msg = str(exc).lower()
        non_recoverable_patterns = [
            "not found",
            "not installed",
            "not supported",
            "invalid",
            "missing",
            "does not exist",
            "no such file",
            "cannot find",
        ]
        for pattern in non_recoverable_patterns:
            if pattern in exc_msg:
                return True

        return False

    def _dispatch(
        self,
        *,
        category: object,
        metadata: Mapping[str, object],
        dry_run: bool,
    ) -> Mapping[str, object] | None:
        if category == "context":
            self._human_logger.info("Context: %s", metadata)
            return None
        if category == "detect":
            summary = metadata.get("summary") if isinstance(metadata, dict) else None
            if summary:
                self._human_logger.info("Detection summary: %s", summary)
            else:
                self._human_logger.info("Detection snapshot captured.")
            return None
        if category == "msi-uninstall":
            product = metadata.get("product")
            if not product:
                self._human_logger.warning(
                    "Skipping MSI uninstall step without product metadata: %s",
                    metadata,
                )
            else:
                force = bool(metadata.get("force", False))
                if force:
                    # Force mode: terminate Office processes before MSI uninstall
                    self._human_logger.info("Force mode: terminating Office processes")
                    processes.terminate_office_processes(constants.DEFAULT_OFFICE_PROCESSES)
                    processes.terminate_process_patterns(constants.OFFICE_PROCESS_PATTERNS)
                msi_uninstall.uninstall_products([product], dry_run=dry_run)  # type: ignore[list-item]
            return None
        if category == "c2r-uninstall":
            installation = metadata.get("installation") or metadata
            if not installation:
                self._human_logger.warning(
                    "Skipping C2R uninstall step without installation metadata",
                )
            else:
                force = bool(metadata.get("force", False))
                c2r_uninstall.uninstall_products(installation, dry_run=dry_run, force=force)
            return None
        if category == "licensing-cleanup":
            extended = dict(metadata)
            extended["dry_run"] = dry_run
            licensing.cleanup_licenses(extended)
            return None
        if category == "task-cleanup":
            tasks = [str(task) for task in metadata.get("tasks", []) if task]
            if not tasks:
                self._human_logger.info("No scheduled tasks supplied; skipping step.")
            else:
                tasks_services.remove_tasks(tasks, dry_run=dry_run)
            return None
        if category == "service-cleanup":
            services = [str(service) for service in metadata.get("services", []) if service]
            if not services:
                self._human_logger.info("No services supplied; skipping step.")
            else:
                tasks_services.delete_services(services, dry_run=dry_run)
            return None
        if category == "filesystem-cleanup":
            _perform_filesystem_cleanup(
                metadata,
                self._context_metadata,
                dry_run=dry_run,
            )
            return None
        if category == "registry-cleanup":
            backup_info = _perform_registry_cleanup(
                metadata,
                dry_run=dry_run,
                default_backup=self._backup_destination,
                default_logdir=self._log_directory,
            )
            return dict(backup_info)

        self._human_logger.info("Unhandled plan category %s; skipping.", category)
        return None

    @staticmethod
    def _resolve_retry_count(step: Mapping[str, object], metadata: Mapping[str, object]) -> int:
        """!
        @brief Resolve retry count from step/metadata, defaulting to DEFAULT_RETRY_COUNT.
        """
        values = [
            step.get("retries"),
            metadata.get("retries"),
            metadata.get("retry_attempts"),
        ]
        for value in values:
            if value is None:
                continue
            try:
                parsed = int(value)
            except (TypeError, ValueError):
                continue
            return max(0, parsed)
        return DEFAULT_RETRY_COUNT

    @staticmethod
    def _resolve_retry_delay(step: Mapping[str, object], metadata: Mapping[str, object]) -> int:
        """!
        @brief Resolve base retry delay from step/metadata, defaulting to DEFAULT_RETRY_DELAY_BASE.
        """
        values = [
            step.get("retry_delay"),
            metadata.get("retry_delay"),
            metadata.get("retry_delay_seconds"),
        ]
        for value in values:
            if value is None:
                continue
            try:
                parsed = int(value)
            except (TypeError, ValueError):
                continue
            return max(0, parsed)
        return DEFAULT_RETRY_DELAY_BASE

    @staticmethod
    def _calculate_progressive_delay(base_delay: int, attempt: int) -> int:
        """!
        @brief Calculate progressive delay with exponential backoff.
        @details Uses formula: base_delay * (1.5 ^ (attempt - 1)), capped at MAX.
        """
        factor = 1.5 ** (attempt - 1)
        calculated = int(base_delay * factor)
        return min(calculated, DEFAULT_RETRY_DELAY_MAX)


def _merge_reboot_details(details: MutableMapping[str, object], services: Iterable[str]) -> None:
    """!
    @brief Merge ``services`` into ``details`` reboot recommendation payload.
    """

    normalized = [text for text in (str(service).strip() for service in services) if text]
    if not normalized:
        return

    current = details.get("reboot_services")
    existing: list[str] = []
    if isinstance(current, Iterable) and not isinstance(current, (str, bytes)):
        existing = [str(item).strip() for item in current if str(item).strip()]

    combined = list(dict.fromkeys([*existing, *normalized]))
    details["reboot_services"] = combined
    details["reboot_recommended"] = True


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
    max_pass_limit = int(
        max_passes
        or options.get("max_passes")
        or context_metadata.get("max_passes", DEFAULT_MAX_PASSES)
        or DEFAULT_MAX_PASSES
    )

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

    # Pre-scrub process/service cleanup
    if global_dry_run:
        _scrub_progress("DRY RUN MODE - No destructive actions will occur")
        human_logger.info("Executing plan in dry-run mode; no destructive actions will occur.")
    else:
        _scrub_progress("Terminating Office processes...", newline=False)
        processes.terminate_office_processes(constants.DEFAULT_OFFICE_PROCESSES)
        processes.terminate_process_patterns(constants.OFFICE_PROCESS_PATTERNS)
        _scrub_ok()

        _scrub_progress("Stopping Office services...", newline=False)
        tasks_services.stop_services(constants.KNOWN_SERVICES)
        _scrub_ok()

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
            f"Pass {current_pass} complete: {pass_successes} success, {pass_failures} failed, {pass_skipped} skipped ({pass_duration:.2f}s)",
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
            human_logger.warning(
                "Reached maximum scrub passes (%d); continuing to cleanup phase.",
                max_pass_limit,
            )
            final_plan = current_plan
            break

        _scrub_progress("Re-probing inventory for next pass...", indent=1)
        inventory = detect.reprobe(base_options)
        next_plan_raw = plan_module.build_plan(inventory, base_options, pass_index=current_pass + 1)
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
        _scrub_progress(
            f"Cleanup steps to process: {sum(1 for s in final_plan if s.get('category') in CLEANUP_CATEGORIES)}"
        )

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
        try:
            cleanup_results = _execute_steps(final_plan, CLEANUP_CATEGORIES, global_dry_run)
        except StepExecutionError as exc:
            _scrub_progress("Cleanup phase FAILED")
            all_results.extend(exc.partial_results)
            _log_summary(all_results, passes_run, global_dry_run)
            raise
        else:
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
            f"Cleanup complete: {cleanup_successes} success, {cleanup_failures} failed ({cleanup_duration:.2f}s)"
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
) -> list[StepResult]:
    """!
    @brief Execute the subset of plan steps matching ``categories``.
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
            # Non-recoverable errors: log and continue to next step
            # Recoverable errors: stop and allow pass retry
            if result.non_recoverable:
                logging_ext.get_human_logger().warning(
                    "Step %s failed with non-recoverable error, continuing to next step...",
                    result.step_id or result.category,
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
    purge_metadata = metadata.get("purge_templates")
    if purge_metadata is not None:
        purge_templates = bool(purge_metadata)
    else:
        purge_templates = bool(options.get("force", False) and not preserve_templates)

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
) -> Mapping[str, object]:
    """!
    @brief Export and delete registry leftovers with backup awareness.
    @details Consolidates the registry cleanup logic so backup destinations are
    normalised once and deletions are skipped when no keys remain. The helper
    reuses plan metadata when provided and generates a timestamped backup path when
    only a log directory is available, mirroring the OffScrub behaviour. Returns a
    mapping describing whether a backup destination was requested or written so
    the caller can surface the information in the final summary.
    """

    human_logger = logging_ext.get_human_logger()

    keys = _normalize_string_sequence(metadata.get("keys", []))
    keys = _sort_registry_paths_deepest_first(keys)
    if not keys:
        human_logger.info("No registry keys supplied; skipping step.")
        return {"backup_requested": False, "backup_performed": False, "keys_processed": 0}

    step_backup = _normalize_option_path(metadata.get("backup_destination")) or default_backup
    step_logdir = _normalize_option_path(metadata.get("log_directory")) or default_logdir

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
            human_logger.info(
                "Exporting %d registry keys to %s before deletion.",
                len(keys),
                step_backup,
            )
            registry_tools.export_keys(keys, step_backup)
            backup_performed = True
        else:
            human_logger.warning("Proceeding without registry backup; no destination available.")

    registry_tools.delete_keys(keys, dry_run=dry_run)

    return {
        "backup_destination": step_backup,
        "backup_requested": backup_requested,
        "backup_performed": backup_performed,
        "keys_processed": len(keys),
    }


def _normalize_string_sequence(values: object) -> list[str]:
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


def _sort_registry_paths_deepest_first(paths: Iterable[str]) -> list[str]:
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
