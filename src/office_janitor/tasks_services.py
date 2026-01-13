"""!
@brief Scheduled task and service management utilities.
@details Wraps ``schtasks.exe`` and ``sc.exe`` to disable/delete scheduled
Office tasks, stop/start related services, and poll service state with
retry-aware logging. The helpers mirror OffScrub automation semantics while
respecting dry-run and timeout safeguards.
"""

from __future__ import annotations

import time
from collections.abc import Iterable, Sequence
from contextlib import contextmanager

from . import constants, exec_utils, logging_ext


def disable_tasks(task_names: Iterable[str], *, dry_run: bool = False) -> None:
    """!
    @brief Disable scheduled tasks linked to Office components.
    @details Executes ``schtasks /Change /Disable`` unless ``dry_run`` is set.
    @param task_names Iterable of task paths to disable.
    @param dry_run When ``True`` just logs the intended action.
    """

    human_logger = logging_ext.get_human_logger()

    tasks: list[str] = [name for name in (str(name).strip() for name in task_names) if name]
    for task in tasks:
        result = exec_utils.run_command(
            ["schtasks.exe", "/Change", "/TN", task, "/Disable"],
            event="task_disable",
            timeout=60,
            dry_run=dry_run,
            human_message=f"Disabling scheduled task {task}",
            extra={"task": task},
        )

        if result.skipped:
            continue

        if result.returncode == 127:
            human_logger.debug("schtasks.exe unavailable; cannot disable %s", task)
            continue

        if result.returncode == 0 and not result.error:
            human_logger.info("Disabled scheduled task %s", task)
        else:
            human_logger.debug(
                "schtasks exited with %s for %s: %s",
                result.returncode,
                task,
                result.stderr.strip(),
            )


def delete_tasks(task_names: Sequence[str], *, dry_run: bool = False) -> None:
    """!
    @brief Delete scheduled tasks using ``schtasks /Delete`` semantics.
    @param task_names Tasks to remove.
    @param dry_run When ``True`` skip executing the command.
    """

    human_logger = logging_ext.get_human_logger()

    for task in (str(name).strip() for name in task_names if str(name).strip()):
        result = exec_utils.run_command(
            ["schtasks.exe", "/Delete", "/TN", task, "/F"],
            event="task_delete",
            timeout=60,
            dry_run=dry_run,
            human_message=f"Deleting scheduled task {task}",
            extra={"task": task},
        )

        if result.skipped:
            continue

        if result.returncode == 127:
            human_logger.debug("schtasks.exe unavailable; cannot delete %s", task)
            continue

        if result.returncode == 0 and not result.error:
            human_logger.info("Deleted scheduled task %s", task)
        else:
            human_logger.debug(
                "schtasks exited with %s for %s: %s",
                result.returncode,
                task,
                result.stderr.strip(),
            )


_PENDING_REBOOT_SERVICES: set[str] = set()
"""!
@brief Services that could not be stopped cleanly and require a reboot.
"""

_SUPPRESS_REBOOT_RECOMMENDATIONS = False
"""!
@brief When ``True``, reboot recommendations are not recorded (legacy /NOREBOOT).
"""


def _record_reboot_recommendation(service: str) -> None:
    """!
    @brief Track ``service`` as requiring a reboot to finish shutting down.
    """

    if _SUPPRESS_REBOOT_RECOMMENDATIONS:
        return
    clean_name = str(service).strip()
    if not clean_name:
        return
    _PENDING_REBOOT_SERVICES.add(clean_name)


def consume_reboot_recommendations() -> list[str]:
    """!
    @brief Return and clear the accumulated reboot recommendations.
    @details ``stop_services`` records any services that timed out while
    stopping. This helper exposes the aggregated list so scrub summaries can
    remind operators to reboot the host before retrying Office tasks.
    """

    if not _PENDING_REBOOT_SERVICES:
        return []
    services = sorted(_PENDING_REBOOT_SERVICES)
    _PENDING_REBOOT_SERVICES.clear()
    return services


def get_reboot_recommendations() -> list[str]:
    """!
    @brief Return any recorded reboot recommendations without clearing them.
    """

    if not _PENDING_REBOOT_SERVICES:
        return []
    return sorted(_PENDING_REBOOT_SERVICES)


def reboot_recommendations_suppressed() -> bool:
    """!
    @brief Return whether reboot recommendation recording is currently suppressed.
    """

    return bool(_SUPPRESS_REBOOT_RECOMMENDATIONS)


@contextmanager
def suppress_reboot_recommendations(enabled: bool = True):
    """!
    @brief Temporarily suppress recording of reboot recommendations.
    @details Mirrors legacy ``/NOREBOOT`` behaviour by preventing
    :func:`_record_reboot_recommendation` from tracking services during the
    managed block.
    """

    global _SUPPRESS_REBOOT_RECOMMENDATIONS
    previous = _SUPPRESS_REBOOT_RECOMMENDATIONS
    _SUPPRESS_REBOOT_RECOMMENDATIONS = bool(enabled)
    try:
        yield
    finally:
        _SUPPRESS_REBOOT_RECOMMENDATIONS = previous


def stop_services(service_names: Iterable[str], *, timeout: int = 30) -> dict[str, object]:
    """!
    @brief Stop services that keep Office components resident.
    @details Issues ``sc stop`` followed by ``sc config start= disabled`` to
    prevent restarts during cleanup.
    @param service_names Iterable of service names.
    @param timeout Maximum seconds for each subprocess call.
    @returns Dictionary containing ``reboot_required`` and
    ``services_requiring_reboot`` flags for downstream summaries.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    services: list[str] = [name for name in (str(name).strip() for name in service_names) if name]
    reboot_services: list[str] = []

    for service in services:
        stop_result = exec_utils.run_command(
            ["sc.exe", "stop", service],
            event="service_stop",
            timeout=timeout,
            human_message=f"Stopping service {service}",
            extra={"service": service},
        )

        if stop_result.returncode == 127:
            human_logger.debug("sc.exe unavailable; cannot stop %s", service)
            continue

        if stop_result.timed_out:
            reboot_services.append(service)
            _record_reboot_recommendation(service)
            human_logger.warning(
                "Timed out stopping service %s; recommend reboot to finish shutting it down.",
                service,
            )
            machine_logger.warning(
                "service_stop_timeout",
                extra={
                    "event": "service_stop_timeout",
                    "service": service,
                    "reboot_required": True,
                },
            )

        if stop_result.returncode == 0 and not stop_result.error:
            human_logger.info("Stopped service %s", service)
        else:
            human_logger.debug("Service %s stop returned %s", service, stop_result.returncode)

        disable_result = exec_utils.run_command(
            ["sc.exe", "config", service, "start=", "disabled"],
            event="service_disable",
            timeout=timeout,
            human_message=f"Disabling service {service}",
            extra={"service": service},
        )

        if disable_result.returncode == 127:
            human_logger.debug("sc.exe unavailable; cannot disable %s", service)
            continue

        if disable_result.timed_out:
            human_logger.warning("Timed out disabling service %s", service)
            continue

        if disable_result.returncode == 0 and not disable_result.error:
            human_logger.info("Configured service %s to be disabled", service)
        else:
            human_logger.debug("Service %s disable returned %s", service, disable_result.returncode)

    unique_reboot_services = list(dict.fromkeys(reboot_services))

    return {
        "reboot_required": bool(unique_reboot_services),
        "services_requiring_reboot": unique_reboot_services,
    }


def start_services(service_names: Iterable[str], *, timeout: int = 30) -> None:
    """!
    @brief Start services previously stopped for cleanup.
    @param service_names Iterable of service names.
    @param timeout Maximum seconds for each ``sc start`` invocation.
    """

    human_logger = logging_ext.get_human_logger()

    services: list[str] = [name for name in (str(name).strip() for name in service_names) if name]
    for service in services:
        result = exec_utils.run_command(
            ["sc.exe", "start", service],
            event="service_start",
            timeout=timeout,
            human_message=f"Starting service {service}",
            extra={"service": service},
        )

        if result.returncode == 127:
            human_logger.debug("sc.exe unavailable; cannot start %s", service)
            continue

        if result.timed_out:
            human_logger.warning("Timed out starting service %s", service)
            continue

        if result.returncode == 0 and not result.error:
            human_logger.info("Started service %s", service)
        else:
            human_logger.debug("Service %s start returned %s", service, result.returncode)


def delete_services(service_names: Sequence[str], *, dry_run: bool = False) -> None:
    """!
    @brief Remove services entirely using ``sc delete``.
    @param service_names Sequence of services to delete.
    @param dry_run When ``True`` only log the intended action.
    """

    human_logger = logging_ext.get_human_logger()

    for service in (str(name).strip() for name in service_names if str(name).strip()):
        result = exec_utils.run_command(
            ["sc.exe", "delete", service],
            event="service_delete",
            timeout=30,
            dry_run=dry_run,
            human_message=f"Deleting service {service}",
            extra={"service": service},
        )

        if result.skipped:
            continue

        if result.returncode == 127:
            human_logger.debug("sc.exe unavailable; cannot delete %s", service)
            continue

        if result.returncode == 0 and not result.error:
            human_logger.info("Deleted service %s", service)
        else:
            human_logger.debug("Service %s delete returned %s", service, result.returncode)


def query_service_status(
    service: str,
    *,
    retries: int = 3,
    delay: float = 1.0,
    timeout: int = 30,
) -> str:
    """!
    @brief Query the current status of a Windows service with retry support.
    @details Runs ``sc query`` up to ``retries`` times, waiting ``delay``
    seconds between attempts if the command errors or times out. The final
    recognised status string (``RUNNING``, ``STOPPED``, etc.) is returned in
    uppercase. If all attempts fail ``"UNKNOWN"`` is returned.
    @param service Service name to query.
    @param retries Number of attempts before giving up.
    @param delay Seconds to wait between attempts.
    @param timeout Timeout for each ``sc query`` execution.
    @returns Service status string.
    """

    service_name = str(service).strip()
    if not service_name:
        return "UNKNOWN"

    human_logger = logging_ext.get_human_logger()

    for attempt in range(1, max(1, retries) + 1):
        result = exec_utils.run_command(
            ["sc.exe", "query", service_name],
            event="service_query",
            timeout=timeout,
            extra={"service": service_name, "attempt": attempt},
        )

        if result.returncode == 127:
            human_logger.debug("sc.exe unavailable; cannot query %s", service_name)
            return "UNKNOWN"

        if result.timed_out:
            human_logger.warning(
                "Timed out querying status for %s (attempt %d)", service_name, attempt
            )
            if attempt < retries:
                time.sleep(delay)
            continue

        if result.returncode == 0 and not result.error:
            status = _parse_service_state(result.stdout)
            if status:
                human_logger.debug("Service %s status: %s", service_name, status)
                return status
        else:
            human_logger.debug(
                "sc query for %s returned %s",
                service_name,
                result.returncode,
            )

        if attempt < retries:
            time.sleep(delay)

    human_logger.debug("Service %s status unknown after %d attempts", service_name, retries)
    return "UNKNOWN"


def _parse_service_state(output: str) -> str:
    """!
    @brief Extract the status token from ``sc query`` output.
    @param output Raw stdout text from the command.
    @returns Uppercase status token or empty string when not detected.
    """

    for line in output.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.upper().startswith("STATE"):
            _, _, remainder = stripped.partition(":")
            tokens = remainder.strip().split()
            if tokens:
                return tokens[-1].upper()
    return ""


def remove_tasks(task_names: Sequence[str], *, dry_run: bool = False) -> None:
    """!
    @brief Backwards-compatible wrapper for :func:`delete_tasks`.
    @param task_names Tasks to remove.
    @param dry_run When ``True`` skip executing the command.
    """

    delete_tasks(task_names, dry_run=dry_run)


def delete_office_scheduled_tasks(*, dry_run: bool = False) -> None:
    """!
    @brief Delete all known Office scheduled tasks.
    @details Uses the comprehensive ``OFFICE_SCHEDULED_TASKS_TO_DELETE`` constant
    derived from OffScrubC2R.vbs DelSchtasks subroutine to remove streaming,
    telemetry, subscription, and update tasks.
    @param dry_run When ``True`` only log the intended actions.
    """

    delete_tasks(constants.OFFICE_SCHEDULED_TASKS_TO_DELETE, dry_run=dry_run)


def delete_office_services(*, dry_run: bool = False) -> None:
    """!
    @brief Delete all known Office services.
    @details Uses the ``OFFICE_SERVICES_TO_DELETE`` constant to remove
    ClickToRunSvc, OfficeSvc, OSE, and licensing services.
    @param dry_run When ``True`` only log the intended actions.
    """

    delete_services(constants.OFFICE_SERVICES_TO_DELETE, dry_run=dry_run)


# ---------------------------------------------------------------------------
# OSE Service State Validation
# ---------------------------------------------------------------------------
# Based on OffScrubC2R.vbs Uninstall subroutine (lines 1224-1233)


def validate_ose_service_state(
    *,
    dry_run: bool = False,
    timeout: int = 30,
) -> dict[str, object]:
    """!
    @brief Validate and fix OSE service state before uninstall operations.
    @details Implements the VBS OSE service checks:
        1. If OSE is disabled, change to Manual start mode
        2. If OSE is not running as LocalSystem, change the service account
    @param dry_run When ``True`` only log what would be changed.
    @param timeout Timeout for each sc.exe command.
    @returns Dictionary with validation results and any changes made.
    """
    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    results: dict[str, object] = {
        "ose_found": False,
        "ose64_found": False,
        "disabled_fixed": [],
        "account_fixed": [],
        "errors": [],
    }

    ose_services = ["ose", "ose64"]

    for service in ose_services:
        # Query the service configuration
        result = exec_utils.run_command(
            ["sc.exe", "qc", service],
            event="service_query_config",
            timeout=timeout,
            extra={"service": service},
        )

        if result.returncode != 0:
            # Service not found
            continue

        output = result.stdout
        if service == "ose":
            results["ose_found"] = True
        else:
            results["ose64_found"] = True

        # Parse START_TYPE
        is_disabled = "DISABLED" in output.upper()
        start_type_line = ""
        for line in output.splitlines():
            if "START_TYPE" in line.upper():
                start_type_line = line
                break

        if is_disabled or "4  DISABLED" in output:
            human_logger.info(
                "OSE service %s is disabled, changing to Manual...",
                service,
            )
            machine_logger.info(
                "ose_service_disabled",
                extra={
                    "event": "ose_service_disabled",
                    "service": service,
                    "start_type": start_type_line,
                },
            )

            if not dry_run:
                fix_result = exec_utils.run_command(
                    ["sc.exe", "config", service, "start=", "demand"],
                    event="service_config_start",
                    timeout=timeout,
                    extra={"service": service, "start_mode": "demand"},
                )
                if fix_result.returncode == 0:
                    results["disabled_fixed"].append(service)  # type: ignore[union-attr]
                    human_logger.info("Changed %s to Manual start", service)
                else:
                    results["errors"].append(  # type: ignore[union-attr]
                        {"service": service, "operation": "change_start", "code": fix_result.returncode}
                    )
            else:
                results["disabled_fixed"].append(service)  # type: ignore[union-attr]

        # Parse SERVICE_START_NAME (account)
        is_localsystem = "LOCALSYSTEM" in output.upper()
        account_line = ""
        for line in output.splitlines():
            if "SERVICE_START_NAME" in line.upper():
                account_line = line
                is_localsystem = "LOCALSYSTEM" in line.upper()
                break

        if not is_localsystem:
            human_logger.info(
                "OSE service %s not running as LocalSystem, fixing...",
                service,
            )
            machine_logger.info(
                "ose_service_wrong_account",
                extra={
                    "event": "ose_service_wrong_account",
                    "service": service,
                    "account_line": account_line,
                },
            )

            if not dry_run:
                # sc config ose obj= LocalSystem
                fix_result = exec_utils.run_command(
                    ["sc.exe", "config", service, "obj=", "LocalSystem", "password=", ""],
                    event="service_config_account",
                    timeout=timeout,
                    extra={"service": service, "account": "LocalSystem"},
                )
                if fix_result.returncode == 0:
                    results["account_fixed"].append(service)  # type: ignore[union-attr]
                    human_logger.info("Changed %s to run as LocalSystem", service)
                else:
                    results["errors"].append(  # type: ignore[union-attr]
                        {"service": service, "operation": "change_account", "code": fix_result.returncode}
                    )
            else:
                results["account_fixed"].append(service)  # type: ignore[union-attr]

    if not results["ose_found"] and not results["ose64_found"]:
        human_logger.debug("No OSE service found (may not be installed)")

    return results
