"""!
@brief Process and service control helpers.
@details The process utilities terminate running Office binaries and pause
background services that block uninstall operations, following the
specification's safety and retry requirements.
"""
from __future__ import annotations

from typing import Iterable


def terminate_office_processes(names: Iterable[str], *, timeout: int = 30) -> None:
    """!
    @brief Stop known Office processes before uninstalling.
    """

    raise NotImplementedError
