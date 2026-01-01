"""!
@brief Backwards-compatible shim for subprocess execution helpers.
@details Historically the project used :mod:`office_janitor.command_runner` to
wrap :func:`subprocess.run`. The functionality now lives in
:mod:`office_janitor.exec_utils`, but this module re-exports the primary
interfaces so existing imports continue to function without modification.
"""

from __future__ import annotations

from .exec_utils import CommandResult, run_command

__all__ = ["CommandResult", "run_command"]
