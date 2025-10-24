"""!
@brief Registry management helpers.
@details The registry tools export hives for backup, delete targeted keys, and
provide winreg utilities used throughout detection and cleanup as outlined in
the specification.
"""
from __future__ import annotations

from typing import Iterable


def export_keys(keys: Iterable[str], destination: str) -> None:
    """!
    @brief Export the provided registry keys to ``.reg`` files in ``destination``.
    """

    raise NotImplementedError


def delete_keys(keys: Iterable[str], *, dry_run: bool = False) -> None:
    """!
    @brief Remove registry keys while respecting dry-run safeguards.
    """

    raise NotImplementedError
