"""!
@brief Version metadata for Office Janitor.
@details The version module centralizes version and build identifiers so that
the command-line interface and user interfaces present consistent information.
"""
from __future__ import annotations

from typing import Dict

__all__ = ["__version__", "__build__", "build_info"]

__version__ = "0.0.0"
__build__ = "dev"


def build_info() -> Dict[str, str]:
    """!
    @brief Provide a mapping with the current version metadata.
    @returns Dictionary containing ``version`` and ``build`` keys.
    """

    return {"version": __version__, "build": __build__}
