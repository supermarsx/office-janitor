"""!
@brief Shim entry point for Office Janitor.
@details This module ensures the package in ``src/`` is importable before
transferring control to :func:`office_janitor.main.main`.
"""
from __future__ import annotations

import os
import sys


def _prepend_src_to_sys_path() -> None:
    """!
    @brief Prepend the repository ``src`` directory to ``sys.path``.
    @details The shim mirrors the structure described in :mod:`spec.md`, keeping
    the distributable executable simple while letting the package live under
    ``src/``.
    """

    repo_root = os.path.dirname(__file__)
    src_path = os.path.join(repo_root, "src")
    if src_path not in sys.path:
        sys.path.insert(0, src_path)


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
