"""!
@brief Shim entry point for Office Janitor.
@details This module ensures the package in ``src/`` is importable before
transferring control to :func:`office_janitor.main.main`.
"""
from __future__ import annotations

import os
import sys

__all__ = ["main"]

_REPO_ROOT = os.path.dirname(__file__)
_SRC_PATH = os.path.join(_REPO_ROOT, "src")
_PACKAGE_PATH = os.path.join(_SRC_PATH, "office_janitor")

if os.path.isdir(_SRC_PATH) and _SRC_PATH not in sys.path:
    sys.path.insert(0, _SRC_PATH)

if os.path.isdir(_PACKAGE_PATH):
    __path__ = [_PACKAGE_PATH]
    if __spec__ is not None:  # pragma: no cover - import system attribute
        __spec__.submodule_search_locations = list(__path__)


def _prepend_src_to_sys_path() -> None:
    """!
    @brief Prepend the repository ``src`` directory to ``sys.path``.
    @details The shim mirrors the structure described in :mod:`spec.md`, keeping
    the distributable executable simple while letting the package live under
    ``src/``.
    """

    if _SRC_PATH not in sys.path:
        sys.path.insert(0, _SRC_PATH)


def main() -> int:
    """!
    @brief Invoke the package entry point after preparing ``sys.path``.
    @returns Exit status propagated from :func:`office_janitor.main.main`.
    """

    _prepend_src_to_sys_path()
    from office_janitor.main import main as package_main

    return package_main()


if __name__ == "__main__":  # pragma: no cover - manual invocation
    sys.exit(main())
