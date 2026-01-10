"""!
@brief Shared application state typing helpers for UI layers.
@details Defines the mapping structure exchanged between :mod:`ui`, :mod:`tui`,
and :mod:`main` so MyPy can validate cross-module access to loggers, event
queues, and callable hooks without resorting to ``Any``.
"""

from __future__ import annotations

import argparse
import logging
from collections import deque
from collections.abc import Callable, Mapping
from typing import Deque, TypedDict


class _RequiredAppState(TypedDict):
    args: argparse.Namespace
    human_logger: logging.Logger
    machine_logger: logging.Logger
    detector: Callable[[], Mapping[str, object]]
    planner: Callable[[Mapping[str, object], Mapping[str, object] | None], list[dict[str, object]]]
    executor: Callable[[list[dict[str, object]], Mapping[str, object] | None], bool | None]
    event_queue: Deque[dict[str, object]]
    emit_event: Callable[..., None]
    confirm: Callable[..., bool]


class AppState(_RequiredAppState, total=False):
    input: Callable[[str], str]


def new_event_queue() -> Deque[dict[str, object]]:
    """!
    @brief Provide a typed ``deque`` for UI events.
    """

    return deque()

