"""!
@brief Scheduled task and service management utilities.
@details Implements discovery and cleanup of scheduled tasks, services, and
related artifacts that keep Office components resident, matching the
specification's guidelines.
"""
from __future__ import annotations

from typing import Iterable


def disable_tasks(task_names: Iterable[str], *, dry_run: bool = False) -> None:
    """!
    @brief Disable or remove scheduled tasks linked to Office.
    """

    raise NotImplementedError


def stop_services(service_names: Iterable[str], *, timeout: int = 30) -> None:
    """!
    @brief Stop service processes before uninstall operations proceed.
    """

    raise NotImplementedError
