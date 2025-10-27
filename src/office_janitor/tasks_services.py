"""!
@brief Scheduled task and service management utilities.
@details Wraps ``schtasks.exe`` and ``sc.exe`` to disable/delete scheduled
Office tasks, stop/start related services, and poll service state with
retry-aware logging. The helpers mirror OffScrub automation semantics while
respecting dry-run and timeout safeguards.
"""
from __future__ import annotations

import subprocess
import time
from typing import Iterable, List, Sequence

from . import logging_ext


def disable_tasks(task_names: Iterable[str], *, dry_run: bool = False) -> None:
    """!
    @brief Disable scheduled tasks linked to Office components.
    @details Executes ``schtasks /Change /Disable`` unless ``dry_run`` is set.
    @param task_names Iterable of task paths to disable.
    @param dry_run When ``True`` just logs the intended action.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    tasks: List[str] = [name for name in (str(name).strip() for name in task_names) if name]
    for task in tasks:
        command = ["schtasks.exe", "/Change", "/TN", task, "/Disable"]
        machine_logger.info(
            "task_disable_plan",
            extra={
                "event": "task_disable_plan",
                "task": task,
                "dry_run": bool(dry_run),
                "command": command,
            },
        )

        if dry_run:
            human_logger.info("Dry-run: would disable scheduled task %s", task)
            continue

        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=60,
                check=False,
            )
        except FileNotFoundError:  # pragma: no cover - non-Windows fallback
            human_logger.debug("schtasks.exe unavailable; cannot disable %s", task)
            continue

        machine_logger.info(
            "task_disable_result",
            extra={
                "event": "task_disable_result",
                "task": task,
                "return_code": result.returncode,
                "stdout": result.stdout,
                "stderr": result.stderr,
            },
        )
        if result.returncode == 0:
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
    machine_logger = logging_ext.get_machine_logger()

    for task in (str(name).strip() for name in task_names if str(name).strip()):
        command = ["schtasks.exe", "/Delete", "/TN", task, "/F"]
        machine_logger.info(
            "task_delete_plan",
            extra={
                "event": "task_delete_plan",
                "task": task,
                "dry_run": bool(dry_run),
                "command": command,
            },
        )
        if dry_run:
            human_logger.info("Dry-run: would delete scheduled task %s", task)
            continue

        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=60,
                check=False,
            )
        except FileNotFoundError:  # pragma: no cover - non-Windows fallback
            human_logger.debug("schtasks.exe unavailable; cannot delete %s", task)
            continue

        machine_logger.info(
            "task_delete_result",
            extra={
                "event": "task_delete_result",
                "task": task,
                "return_code": result.returncode,
                "stdout": result.stdout,
                "stderr": result.stderr,
            },
        )
        if result.returncode == 0:
            human_logger.info("Deleted scheduled task %s", task)
        else:
            human_logger.debug(
                "schtasks exited with %s for %s: %s",
                result.returncode,
                task,
                result.stderr.strip(),
            )


def stop_services(service_names: Iterable[str], *, timeout: int = 30) -> None:
    """!
    @brief Stop services that keep Office components resident.
    @details Issues ``sc stop`` followed by ``sc config start= disabled`` to
    prevent restarts during cleanup.
    @param service_names Iterable of service names.
    @param timeout Maximum seconds for each subprocess call.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    services: List[str] = [name for name in (str(name).strip() for name in service_names) if name]
    for service in services:
        stop_command = ["sc.exe", "stop", service]
        disable_command = ["sc.exe", "config", service, "start=", "disabled"]
        machine_logger.info(
            "service_stop_plan",
            extra={
                "event": "service_stop_plan",
                "service": service,
                "command": stop_command,
            },
        )

        try:
            stop_result = subprocess.run(
                stop_command,
                capture_output=True,
                text=True,
                timeout=timeout,
                check=False,
            )
        except FileNotFoundError:  # pragma: no cover - non-Windows fallback
            human_logger.debug("sc.exe unavailable; cannot stop %s", service)
            continue
        except subprocess.TimeoutExpired as exc:
            human_logger.warning("Timed out stopping service %s", service)
            machine_logger.warning(
                "service_stop_timeout",
                extra={
                    "event": "service_stop_timeout",
                    "service": service,
                    "stdout": exc.stdout,
                    "stderr": exc.stderr,
                },
            )
            continue

        machine_logger.info(
            "service_stop_result",
            extra={
                "event": "service_stop_result",
                "service": service,
                "return_code": stop_result.returncode,
                "stdout": stop_result.stdout,
                "stderr": stop_result.stderr,
            },
        )
        if stop_result.returncode == 0:
            human_logger.info("Stopped service %s", service)
        else:
            human_logger.debug(
                "Service %s stop returned %s", service, stop_result.returncode
            )

        machine_logger.info(
            "service_disable_plan",
            extra={
                "event": "service_disable_plan",
                "service": service,
                "command": disable_command,
            },
        )
        try:
            disable_result = subprocess.run(
                disable_command,
                capture_output=True,
                text=True,
                timeout=timeout,
                check=False,
            )
        except FileNotFoundError:  # pragma: no cover - non-Windows fallback
            human_logger.debug("sc.exe unavailable; cannot disable %s", service)
            continue
        except subprocess.TimeoutExpired as exc:
            human_logger.warning("Timed out disabling service %s", service)
            machine_logger.warning(
                "service_disable_timeout",
                extra={
                    "event": "service_disable_timeout",
                    "service": service,
                    "stdout": exc.stdout,
                    "stderr": exc.stderr,
                },
            )
            continue

        machine_logger.info(
            "service_disable_result",
            extra={
                "event": "service_disable_result",
                "service": service,
                "return_code": disable_result.returncode,
                "stdout": disable_result.stdout,
                "stderr": disable_result.stderr,
            },
        )
        if disable_result.returncode == 0:
            human_logger.info("Configured service %s to be disabled", service)
        else:
            human_logger.debug(
                "Service %s disable returned %s", service, disable_result.returncode
            )


def start_services(service_names: Iterable[str], *, timeout: int = 30) -> None:
    """!
    @brief Start services previously stopped for cleanup.
    @param service_names Iterable of service names.
    @param timeout Maximum seconds for each ``sc start`` invocation.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    services: List[str] = [name for name in (str(name).strip() for name in service_names) if name]
    for service in services:
        command = ["sc.exe", "start", service]
        machine_logger.info(
            "service_start_plan",
            extra={
                "event": "service_start_plan",
                "service": service,
                "command": command,
            },
        )
        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=timeout,
                check=False,
            )
        except FileNotFoundError:  # pragma: no cover - non-Windows fallback
            human_logger.debug("sc.exe unavailable; cannot start %s", service)
            continue
        except subprocess.TimeoutExpired as exc:
            human_logger.warning("Timed out starting service %s", service)
            machine_logger.warning(
                "service_start_timeout",
                extra={
                    "event": "service_start_timeout",
                    "service": service,
                    "stdout": exc.stdout,
                    "stderr": exc.stderr,
                },
            )
            continue

        machine_logger.info(
            "service_start_result",
            extra={
                "event": "service_start_result",
                "service": service,
                "return_code": result.returncode,
                "stdout": result.stdout,
                "stderr": result.stderr,
            },
        )
        if result.returncode == 0:
            human_logger.info("Started service %s", service)
        else:
            human_logger.debug(
                "Service %s start returned %s", service, result.returncode
            )


def delete_services(service_names: Sequence[str], *, dry_run: bool = False) -> None:
    """!
    @brief Remove services entirely using ``sc delete``.
    @param service_names Sequence of services to delete.
    @param dry_run When ``True`` only log the intended action.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    for service in (str(name).strip() for name in service_names if str(name).strip()):
        command = ["sc.exe", "delete", service]
        machine_logger.info(
            "service_delete_plan",
            extra={
                "event": "service_delete_plan",
                "service": service,
                "dry_run": bool(dry_run),
                "command": command,
            },
        )

        if dry_run:
            human_logger.info("Dry-run: would delete service %s", service)
            continue

        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=30,
                check=False,
            )
        except FileNotFoundError:  # pragma: no cover - non-Windows fallback
            human_logger.debug("sc.exe unavailable; cannot delete %s", service)
            continue

        machine_logger.info(
            "service_delete_result",
            extra={
                "event": "service_delete_result",
                "service": service,
                "return_code": result.returncode,
                "stdout": result.stdout,
                "stderr": result.stderr,
            },
        )
        if result.returncode == 0:
            human_logger.info("Deleted service %s", service)
        else:
            human_logger.debug(
                "Service %s delete returned %s", service, result.returncode
            )


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
    machine_logger = logging_ext.get_machine_logger()

    for attempt in range(1, max(1, retries) + 1):
        command = ["sc.exe", "query", service_name]
        machine_logger.info(
            "service_query_plan",
            extra={
                "event": "service_query_plan",
                "service": service_name,
                "attempt": attempt,
                "command": command,
            },
        )
        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=timeout,
                check=False,
            )
        except FileNotFoundError:  # pragma: no cover - non-Windows fallback
            human_logger.debug("sc.exe unavailable; cannot query %s", service_name)
            return "UNKNOWN"
        except subprocess.TimeoutExpired as exc:
            human_logger.warning(
                "Timed out querying status for %s (attempt %d)", service_name, attempt
            )
            machine_logger.warning(
                "service_query_timeout",
                extra={
                    "event": "service_query_timeout",
                    "service": service_name,
                    "attempt": attempt,
                    "stdout": exc.stdout,
                    "stderr": exc.stderr,
                },
            )
            if attempt < retries:
                time.sleep(delay)
            continue

        machine_logger.info(
            "service_query_result",
            extra={
                "event": "service_query_result",
                "service": service_name,
                "attempt": attempt,
                "return_code": result.returncode,
                "stdout": result.stdout,
                "stderr": result.stderr,
            },
        )
        if result.returncode == 0:
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
