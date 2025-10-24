"""!
@brief Plain console user interface helpers.
@details Provides the interactive menu experience described in the specification
for environments that do not support the richer TUI renderer.
"""
from __future__ import annotations

from typing import Mapping


def run_cli(app_state: Mapping[str, object]) -> None:
    """!
    @brief Launch the basic interactive console menu.
    """

    raise NotImplementedError
