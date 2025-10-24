"""!
@brief Scheduled task and service management utilities.
@details Implements discovery and cleanup of scheduled tasks, services, and
related artifacts that keep Office components resident, matching the
specification's guidelines.
"""
from __future__ import annotations

import subprocess
from typing import Iterable, List

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

        disable_result = subprocess.run(
            disable_command,
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
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
