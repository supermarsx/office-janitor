"""!
@brief Scheduled task and service management utilities.
@details Implements discovery and cleanup of scheduled tasks, services, and
related artifacts that keep Office components resident, matching the
specification's guidelines.
"""
from __future__ import annotations

import subprocess
from typing import Iterable, List, Sequence

from . import logging_ext


def disable_tasks(task_names: Iterable[str], *, dry_run: bool = False) -> None:
    """!
    @brief Disable or remove scheduled tasks linked to Office.
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

        if result.returncode != 0:
            human_logger.debug(
                "schtasks exited with %s for %s: %s", result.returncode, task, result.stderr.strip()
            )
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
        else:
            human_logger.info("Disabled scheduled task %s", task)
            machine_logger.info(
                "task_disable_success",
                extra={
                    "event": "task_disable_success",
                    "task": task,
                    "return_code": result.returncode,
                    "stdout": result.stdout,
                    "stderr": result.stderr,
                },
            )


def remove_tasks(task_names: Sequence[str], *, dry_run: bool = False) -> None:
    """!
    @brief Delete scheduled tasks using ``schtasks /Delete`` semantics.
    @details OffScrub removes Office maintenance tasks outright; this helper
    mirrors that behaviour with dry-run support and structured logging.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    for task in (name for name in task_names if str(name).strip()):
        task_name = str(task).strip()
        machine_logger.info(
            "task_delete_plan",
            extra={
                "event": "task_delete_plan",
                "task": task_name,
                "dry_run": bool(dry_run),
            },
        )
        if dry_run:
            human_logger.info("Dry-run: would delete scheduled task %s", task_name)
            continue

        command = ["schtasks.exe", "/Delete", "/TN", task_name, "/F"]
        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=60,
                check=False,
            )
        except FileNotFoundError:  # pragma: no cover - non-Windows fallback
            human_logger.debug("schtasks.exe unavailable; cannot delete %s", task_name)
            continue

        machine_logger.info(
            "task_delete_result",
            extra={
                "event": "task_delete_result",
                "task": task_name,
                "return_code": result.returncode,
                "stdout": result.stdout,
                "stderr": result.stderr,
            },
        )
        if result.returncode == 0:
            human_logger.info("Deleted scheduled task %s", task_name)
        else:
            human_logger.debug("Scheduled task %s delete returned %s", task_name, result.returncode)


def stop_services(service_names: Iterable[str], *, timeout: int = 30) -> None:
    """!
    @brief Stop service processes before uninstall operations proceed.
    """
    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    services: List[str] = [name for name in (str(name).strip() for name in service_names) if name]
    for service in services:
        stop_command = ["sc.exe", "stop", service]
        disable_command = ["sc.exe", "config", service, "start=", "disabled"]
        machine_logger.info(
            "service_stop_plan",
            extra={"event": "service_stop_plan", "service": service},
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


def delete_services(service_names: Sequence[str], *, dry_run: bool = False) -> None:
    """!
    @brief Remove services entirely using ``sc delete``.
    @details Complements :func:`stop_services` by purging obsolete Office
    services after they have been stopped.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    for service in (name for name in service_names if str(name).strip()):
        service_name = str(service).strip()
        machine_logger.info(
            "service_delete_plan",
            extra={
                "event": "service_delete_plan",
                "service": service_name,
                "dry_run": bool(dry_run),
            },
        )

        if dry_run:
            human_logger.info("Dry-run: would delete service %s", service_name)
            continue

        command = ["sc.exe", "delete", service_name]
        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=30,
                check=False,
            )
        except FileNotFoundError:  # pragma: no cover - non-Windows fallback
            human_logger.debug("sc.exe unavailable; cannot delete %s", service_name)
            continue

        machine_logger.info(
            "service_delete_result",
            extra={
                "event": "service_delete_result",
                "service": service_name,
                "return_code": result.returncode,
                "stdout": result.stdout,
                "stderr": result.stderr,
            },
        )
        if result.returncode == 0:
            human_logger.info("Deleted service %s", service_name)
        else:
            human_logger.debug("Service %s delete returned %s", service_name, result.returncode)
