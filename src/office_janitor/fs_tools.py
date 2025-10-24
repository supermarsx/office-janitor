"""!
@brief Filesystem utilities for Office residue cleanup.
@details Future implementations discover install footprints, reset ACLs, and
remove leftovers from program directories and user profiles, matching the
specification requirements.
"""
from __future__ import annotations

from pathlib import Path
from typing import Iterable


def remove_paths(paths: Iterable[Path], *, dry_run: bool = False) -> None:
    """!
    @brief Delete the supplied paths recursively while respecting dry-run behavior.
    """

    raise NotImplementedError


def reset_acl(path: Path) -> None:
    """!
    @brief Reset permissions on ``path`` so cleanup operations can proceed.
    """

    raise NotImplementedError
