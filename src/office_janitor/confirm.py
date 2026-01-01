"""!
@brief Shared confirmation helpers for destructive operations.
@details Provides reusable prompting logic so both the CLI and future UI
implementations ask the same question prior to running a scrub that modifies
the system.
"""

from __future__ import annotations

import sys
from typing import Callable

CONFIRM_PROMPT = (
    "This will remove Microsoft Office and related licensing artifacts "
    "from this machine. Continue? (Y/n)"
)


def request_scrub_confirmation(
    *,
    dry_run: bool,
    force: bool,
    input_func: Callable[[str], str] | None = None,
    interactive: bool | None = None,
) -> bool:
    """!
    @brief Ask the user to confirm a scrub execution.
    @details The helper enforces the specification's confirmation prompt while
    honouring automation scenarios. Dry-run executions and forced runs bypass
    the prompt entirely. Non-interactive contexts similarly default to
    acceptance so unattended invocations are not blocked.
    @param dry_run Whether the pending execution is a dry-run.
    @param force Whether the caller supplied ``--force``.
    @param input_func Optional input function override for interactive UIs.
    @param interactive Optional override to signal if the environment is
    interactive.
    @returns ``True`` when the scrub should proceed.
    """

    if dry_run or force:
        return True

    if interactive is None:
        stdin = getattr(sys, "stdin", None)
        isatty = getattr(stdin, "isatty", None)
        interactive = bool(isatty and isatty())

    if not interactive:
        return True

    if input_func is None:
        input_func = input

    try:
        response = input_func(f"{CONFIRM_PROMPT} ")
    except EOFError:
        return False

    normalized = response.strip().lower()
    return normalized in ("", "y", "yes")


__all__ = ["CONFIRM_PROMPT", "request_scrub_confirmation"]
