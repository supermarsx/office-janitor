"""!
@brief Structured logging helpers for Office Janitor.
@details When implemented this module configures human-readable logs alongside
JSONL telemetry streams, optionally mirroring machine events to stdout as
detailed in the specification.
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Tuple


def setup_logging(root_dir: Path, *, json_to_stdout: bool = False, level: int = logging.INFO) -> Tuple[logging.Logger, logging.Logger]:
    """!
    @brief Set up human and machine loggers.
    @details The function returns a pair of ``logging.Logger`` objects
    representing the human-readable and structured event channels respectively.
    """

    raise NotImplementedError
