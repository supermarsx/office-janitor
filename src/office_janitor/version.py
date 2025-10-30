"""!
@brief Version metadata for Office Janitor.
@details The version module centralizes version and build identifiers so that
the command-line interface and user interfaces present consistent information.
"""
from __future__ import annotations

from importlib import resources
from typing import Dict

__all__ = ["__version__", "__build__", "build_info"]


def _load_version() -> str:
    """!
    @brief Retrieve the package version string from the packaged ``VERSION`` file.
    @details The PyInstaller and packaging workflows share the same version source so
    that dynamic metadata evaluation does not depend on importing the package during
    a build.
    @returns Semantic version string stored in ``VERSION``.
    """

    version_path = resources.files(__package__).joinpath("VERSION")
    try:
        return version_path.read_text(encoding="utf-8").strip()
    except FileNotFoundError:  # pragma: no cover - defensive fallback
        return "0.0.0"


__version__ = _load_version()
__build__ = "dev"


def build_info() -> Dict[str, str]:
    """!
    @brief Provide a mapping with the current version metadata.
    @returns Dictionary containing ``version`` and ``build`` keys.
    """

    return {"version": __version__, "build": __build__}
