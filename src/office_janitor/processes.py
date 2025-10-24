"""!
@brief Process and service control helpers.
@details The process utilities terminate running Office binaries and pause
background services that block uninstall operations, following the
specification's safety and retry requirements.
"""
from __future__ import annotations

import subprocess
from typing import Iterable, List

from . import logging_ext


def terminate_office_processes(names: Iterable[str], *, timeout: int = 30) -> None:
    """!
    @brief Stop known Office processes before uninstalling.
    """
    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    processes: List[str] = [name for name in (str(name).strip() for name in names) if name]
    if not processes:
        human_logger.debug("No Office processes supplied for termination.")
        return

    human_logger.info("Requesting termination of %d Office processes.", len(processes))
    for process in processes:
        command = ["taskkill.exe", "/IM", process, "/F", "/T"]
        machine_logger.info(
            "terminate_process_plan",
            extra={
                "event": "terminate_process_plan",
                "process_name": process,
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
        except FileNotFoundError:  # pragma: no cover - non-Windows fallback.
            human_logger.debug("taskkill.exe is unavailable; skipping termination for %s", process)
            continue
        except subprocess.TimeoutExpired as exc:
            human_logger.warning("Timed out attempting to stop %s", process)
            machine_logger.warning(
                "terminate_process_timeout",
                extra={
                    "event": "terminate_process_timeout",
                    "process_name": process,
                    "stdout": exc.stdout,
                    "stderr": exc.stderr,
                },
            )
            continue

        if result.returncode != 0:
            human_logger.debug(
                "taskkill exited with %s for %s: %s", result.returncode, process, result.stderr.strip()
            )
            machine_logger.info(
                "terminate_process_result",
                extra={
                    "event": "terminate_process_result",
                    "process_name": process,
                    "return_code": result.returncode,
                    "stdout": result.stdout,
                    "stderr": result.stderr,
                },
            )
        else:
            human_logger.info("Terminated %s", process)
            machine_logger.info(
                "terminate_process_success",
                extra={
                    "event": "terminate_process_success",
                    "process_name": process,
                    "return_code": result.returncode,
                    "stdout": result.stdout,
                    "stderr": result.stderr,
                },
            )
