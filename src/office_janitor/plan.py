"""!
@brief Translate detection results into actionable scrub plans.
@details Planning resolves requested modes, target Office versions, and
user-selected options into an ordered sequence of steps for uninstall, cleanup,
and backups, matching the workflow outlined in the specification.
"""
from __future__ import annotations

from typing import Dict, List, Sequence


def build_plan(inventory: Dict[str, Sequence[dict]], options: Dict[str, object]) -> List[dict]:
    """!
    @brief Produce an ordered plan of actions using the current inventory and CLI options.
    """

    raise NotImplementedError
