"""!
@brief Helpers for orchestrating MSI-based Office uninstalls.
@details This module locates MSI product codes, drives ``msiexec`` with the
correct flags, monitors progress, and captures logs according to the
specification.
"""
from __future__ import annotations

from typing import Iterable


def uninstall_products(product_codes: Iterable[str], *, dry_run: bool = False) -> None:
    """!
    @brief Uninstall the supplied MSI product codes while respecting dry-run semantics.
    """

    raise NotImplementedError
