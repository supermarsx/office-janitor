"""!
@file scrub_executor.py
@brief Step execution engine for Office Janitor scrub operations.

@details Provides the StepResult dataclass for tracking step outcomes, the
StepExecutionError exception for propagating failures with partial results,
and the StepExecutor class that handles retry logic, logging, and dispatching
to the appropriate uninstall/cleanup handlers. This module isolates the
step-level execution mechanics from the higher-level pass orchestration in
scrub.py.
"""

from __future__ import annotations

import dataclasses
import time
from collections.abc import Iterable, Mapping, MutableMapping

from . import (
    c2r_uninstall,
    constants,
    licensing,
    logging_ext,
    msi_uninstall,
    processes,
    registry_tools,
    spinner,
    tasks_services,
)

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

DEFAULT_RETRY_COUNT: int = 9
"""Default number of retry attempts for failed steps."""

DEFAULT_RETRY_DELAY_BASE: int = 3
"""Default base delay in seconds between retries."""

DEFAULT_RETRY_DELAY_MAX: int = 30
"""Maximum delay in seconds for exponential backoff."""

FORCE_ESCALATION_ATTEMPT: int = 3
"""Attempt number at which to escalate to force mode."""


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------


@dataclasses.dataclass
class StepResult:
    """!
    @brief Outcome of a single plan step execution attempt.
    @details Captures identifying information (``step_id``, ``category``), execution
    status (``status``, ``attempts``, ``error``), timing data (``started_at``,
    ``completed_at``), and optional details (``details``).
    """

    step_id: str | None = None
    category: str = "unknown"
    status: str = "pending"
    attempts: int = 0
    dry_run: bool = False
    error: str | None = None
    exception: BaseException | None = None
    progress: float = 0.0
    started_at: float | None = None
    completed_at: float | None = None
    details: MutableMapping[str, object] = dataclasses.field(default_factory=dict)
    non_recoverable: bool = False


class StepExecutionError(Exception):
    """!
    @brief Exception raised when a step fails after exhausting retries.
    @details Contains the failed step result and any partial results from
    steps that completed before the failure.
    """

    def __init__(
        self,
        result: StepResult,
        partial_results: list[StepResult] | None = None,
    ) -> None:
        super().__init__(f"Step {result.step_id or result.category} failed: {result.error}")
        self.result = result
        self.partial_results = partial_results or []


# ---------------------------------------------------------------------------
# Progress output helpers (imported from parent for internal use)
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
# StepExecutor class
# ---------------------------------------------------------------------------


class StepExecutor:
    """!
    @brief Execute individual plan steps with retry support.
    @details Handles dispatching to the appropriate handler based on step category,
    tracks success/failure/retry state, and emits structured logging. Retry behaviour
    uses exponential backoff with a configurable base delay and maximum cap.
    """

    def __init__(
        self,
        *,
        dry_run: bool = False,
        context_metadata: Mapping[str, object] | None = None,
        backup_destination: str | None = None,
        log_directory: str | None = None,
        total_steps: int = 1,
    ) -> None:
        self._dry_run = dry_run
        self._context_metadata = dict(context_metadata or {})
        self._backup_destination = backup_destination
        self._log_directory = log_directory
        self._total_steps = max(1, total_steps)
        self._human_logger = logging_ext.get_human_logger()
        self._machine_logger = logging_ext.get_machine_logger()

        # Per-step tracking for result printing
        self._current_step_label: str = ""
        self._current_step_index: int = 0
        self._current_attempt: int = 1
        self._current_dry_run: bool = dry_run

    def run_step(self, step: Mapping[str, object], *, index: int = 1) -> StepResult:
        """!
        @brief Execute a single plan step with retries when applicable.
        @details Evaluates step metadata to determine retry count and delay,
        then attempts execution up to the configured limit. On success the
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
                    self._emit_result(False, f"{error_reason} (non-recoverable, continuing)")
                    break

                if attempt < attempts_allowed:
                    # Print RETRY status with error reason
                    self._emit_result(False, f"retry: {error_reason}")
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
                self._emit_result(False, error_reason)
                break
            else:
                # Print OK status with duration
                duration = self._format_duration(result)
                duration_str = f"{duration:.2f}s" if duration is not None else ""
                self._emit_result(True, duration_str)
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
        # Update spinner with current step - this keeps spinner animating
        step_label = f"{step_id}" if step_id else f"{category}"
        spinner.set_task(f"{step_label} [{index}/{self._total_steps}]")

        # Store step info for result printing (don't print incomplete line)
        self._current_step_label = step_label
        self._current_step_index = index
        self._current_attempt = attempt
        self._current_dry_run = dry_run

        # Emit warnings for potentially slow operations (first attempt only)
        if attempt == 1:
            slow_step_warnings: dict[str, str] = {
                "filesystem-cleanup": (
                    "Filesystem cleanup may take several minutes depending on data volume"
                ),
                "registry-cleanup": ("Registry cleanup may take a minute while exporting backups"),
                "msi-uninstall": "MSI uninstall may take several minutes per product",
                "c2r-uninstall": "Click-to-Run uninstall may take several minutes",
                "odt-uninstall": "ODT uninstall may take several minutes",
                "offscrub-uninstall": "OffScrub cleanup may take several minutes",
            }
            warning = slow_step_warnings.get(str(category))
            if warning:
                _scrub_progress(f"Note: {warning}", indent=2)

        # Log the step start
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

    def _emit_result(self, success: bool, extra: str = "") -> None:
        """Print the step result with consistent formatting."""
        attempt_info = f" (attempt {self._current_attempt})" if self._current_attempt > 1 else ""
        dry_run_marker = " [DRY-RUN]" if self._current_dry_run else ""
        suffix = f" {extra}" if extra else ""

        status = "[  \033[32mOK\033[0m  ]" if success else "[\033[31mFAILED\033[0m]"
        _scrub_progress(
            f"[{self._current_step_index}/{self._total_steps}] "
            f"{self._current_step_label}{attempt_info}{dry_run_marker}... {status}{suffix}",
            indent=2,
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

        # Check for Access denied patterns
        if "access" in exc_msg.lower() and "denied" in exc_msg.lower():
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
                # Extract detailed product info for logging
                if isinstance(product, dict):
                    product_name = product.get("name") or product.get("display_name") or "Unknown"
                    product_code = product.get("product_code") or product.get("code") or ""
                    product_version = product.get("version") or ""
                else:
                    product_name = str(product)
                    product_code = ""
                    product_version = ""

                _scrub_progress(f"Uninstalling MSI product: {product_name}", indent=3)
                if product_code:
                    _scrub_progress(f"  Product code: {product_code}", indent=3)
                if product_version:
                    _scrub_progress(f"  Version: {product_version}", indent=3)

                force = bool(metadata.get("force", False))
                if force:
                    # Force mode: terminate Office processes before MSI uninstall
                    _scrub_progress("Force mode: terminating Office processes...", indent=3)
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
                # Extract detailed C2R info for logging
                if isinstance(installation, dict):
                    release_id = installation.get("release_id") or "Unknown"
                    display_name = (
                        installation.get("display_name") or installation.get("name") or ""
                    )
                    version = installation.get("version") or ""
                    channel = installation.get("channel") or ""
                    install_path = installation.get("install_path") or ""
                else:
                    release_id = str(installation)
                    display_name = ""
                    version = ""
                    channel = ""
                    install_path = ""

                _scrub_progress(f"Uninstalling Click-to-Run: {release_id}", indent=3)
                if display_name and display_name != release_id:
                    _scrub_progress(f"  Display name: {display_name}", indent=3)
                if version:
                    _scrub_progress(f"  Version: {version}", indent=3)
                if channel:
                    _scrub_progress(f"  Channel: {channel}", indent=3)
                if install_path:
                    _scrub_progress(f"  Install path: {install_path}", indent=3)

                force = bool(metadata.get("force", False))
                if force:
                    _scrub_progress("Force mode enabled for C2R uninstall", indent=3)
                c2r_uninstall.uninstall_products(installation, dry_run=dry_run, force=force)
            return None
        if category == "licensing-cleanup":
            _scrub_progress("Cleaning up Office licenses...", indent=3)
            extended = dict(metadata)
            extended["dry_run"] = dry_run
            licensing.cleanup_licenses(extended)
            return None
        if category == "task-cleanup":
            tasks = [str(task) for task in metadata.get("tasks", []) if task]
            if not tasks:
                self._human_logger.info("No scheduled tasks supplied; skipping step.")
            else:
                _scrub_progress(f"Removing {len(tasks)} scheduled tasks...", indent=3)
                tasks_services.remove_tasks(tasks, dry_run=dry_run)
            return None
        if category == "service-cleanup":
            services = [str(service) for service in metadata.get("services", []) if service]
            if not services:
                self._human_logger.info("No services supplied; skipping step.")
            else:
                _scrub_progress(f"Deleting {len(services)} services...", indent=3)
                tasks_services.delete_services(services, dry_run=dry_run)
            return None
        if category == "filesystem-cleanup":
            # Import here to avoid circular dependency
            from . import scrub_cleanup

            scrub_cleanup.perform_filesystem_cleanup(
                metadata,
                self._context_metadata,
                dry_run=dry_run,
            )
            return None
        if category == "registry-cleanup":
            # Import here to avoid circular dependency
            from . import scrub_cleanup

            backup_info = scrub_cleanup.perform_registry_cleanup(
                metadata,
                dry_run=dry_run,
                default_backup=self._backup_destination,
                default_logdir=self._log_directory,
            )
            return dict(backup_info)
        if category == "vnext-identity-cleanup":
            result = registry_tools.cleanup_vnext_identity_registry(dry_run=dry_run)
            return dict(result)
        if category == "taskband-cleanup":
            include_all_users = bool(metadata.get("include_all_users", False))
            result = registry_tools.cleanup_taskband_registry(
                include_all_users=include_all_users, dry_run=dry_run
            )
            return dict(result)
        if category == "published-components-cleanup":
            result = registry_tools.cleanup_published_components(dry_run=dry_run)
            return dict(result)
        if category == "ose-service-validation":
            result = tasks_services.validate_ose_service_state(dry_run=dry_run)
            return dict(result) if isinstance(result, dict) else {"result": result}

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
            return int(max(0, parsed))
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
            return int(max(0, parsed))
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


# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------


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
