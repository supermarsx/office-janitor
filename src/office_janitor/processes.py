"""!
@brief Process management helpers for Office scrubbing.
@details Provides helpers to discover and terminate running Office
processes. The module mirrors the behaviour from the legacy OffScrub
scripts with structured logging, user prompting, and timeout-aware
subprocess execution so automated runs stay safe.
"""
from __future__ import annotations

import fnmatch
from typing import Callable, Iterable, List, Sequence

from . import exec_utils, logging_ext


def enumerate_processes(patterns: Iterable[str], *, timeout: int = 30) -> List[str]:
    """!
    @brief Enumerate running processes matching ``patterns``.
    @details Invokes ``tasklist`` once and filters results against wildcard
    expressions, returning unique process names. Errors are logged and result
    in an empty list so callers can decide on the fallback behaviour.
    @param patterns Wildcard expressions or explicit process names.
    @param timeout Maximum seconds to wait for ``tasklist`` to finish.
    @returns A list of matching process names in lowercase.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    expanded_patterns: List[str] = [
        str(pattern).strip().lower() for pattern in patterns if str(pattern).strip()
    ]
    if not expanded_patterns:
        return []

    listing = exec_utils.run_command(
        ["tasklist.exe"],
        event="process_enumerate",
        timeout=timeout,
        extra={"patterns": expanded_patterns},
    )

    if listing.returncode == 127:
        human_logger.debug("tasklist.exe unavailable; cannot enumerate processes")
        return []

    if listing.timed_out:
        human_logger.warning("Timed out enumerating processes via tasklist")
        machine_logger.warning(
            "process_enumeration_timeout",
            extra={
                "event": "process_enumeration_timeout",
                "patterns": expanded_patterns,
                "stdout": listing.stdout,
                "stderr": listing.stderr,
            },
        )
        return []

    if listing.returncode != 0 or listing.error:
        human_logger.debug(
            "tasklist exited with %s; skipping enumeration", listing.returncode
        )
        machine_logger.info(
            "process_enumeration_failure",
            extra={
                "event": "process_enumeration_failure",
                "patterns": expanded_patterns,
                "return_code": listing.returncode,
                "stdout": listing.stdout,
                "stderr": listing.stderr,
                "error": listing.error,
            },
        )
        return []

    running: List[str] = []
    for line in listing.stdout.splitlines():
        stripped = line.strip()
        if not stripped or stripped.lower().startswith("image name"):
            continue
        name = stripped.split()[0].lower()
        if name not in running:
            running.append(name)

    matched: List[str] = []
    for pattern in expanded_patterns:
        for name in running:
            if fnmatch.fnmatch(name, pattern) and name not in matched:
                matched.append(name)

    machine_logger.info(
        "process_enumeration_result",
        extra={
            "event": "process_enumeration_result",
            "patterns": expanded_patterns,
            "matches": matched,
        },
    )
    return matched


def prompt_user_to_close(
    processes: Sequence[str],
    *,
    input_func: Callable[[str], str] = input,
    attempts: int = 3,
) -> bool:
    """!
    @brief Prompt the operator to close running Office processes.
    @details Presents a human-readable message listing active processes and
    asks whether the tool should continue with forced termination. The prompt
    is repeated up to ``attempts`` times for unrecognised responses.
    @param processes Sequence of process names that remain active.
    @param input_func Callback used to collect user responses (monkeypatchable).
    @param attempts Maximum attempts before assuming refusal.
    @returns ``True`` when the operator consented to termination; otherwise
    ``False``.
    """

    remaining = [str(name).strip() for name in processes if str(name).strip()]
    if not remaining:
        return True

    human_logger = logging_ext.get_human_logger()
    message = (
        "The following Office processes are still running: {names}.\n"
        "Close them now or Office Janitor can attempt to terminate them automatically."
    ).format(names=", ".join(remaining))
    question = "Proceed with forced termination? [y/N]: "
    human_logger.warning(message)

    for attempt in range(attempts):
        response = input_func(question).strip().lower()
        if response in {"y", "yes"}:
            human_logger.info("Operator approved forced process termination.")
            return True
        if response in {"n", "no", ""}:
            human_logger.info("Operator declined forced process termination.")
            return False
        human_logger.debug("Unrecognised response '%s' (attempt %d)", response, attempt + 1)

    human_logger.info("No approval after %d attempts; skipping termination.", attempts)
    return False


def terminate_office_processes(names: Iterable[str], *, timeout: int = 30) -> None:
    """!
    @brief Forcefully terminate the specified processes.
    @details Issues ``taskkill /F /T`` for each process name. Structured logs
    capture command plans, success, or failure without raising on
    non-critical errors so subsequent cleanup steps can continue.
    @param names Collection of process image names.
    @param timeout Maximum seconds to wait for each ``taskkill`` invocation.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    processes: List[str] = [str(name).strip() for name in names if str(name).strip()]
    if not processes:
        human_logger.debug("No Office processes supplied for termination.")
        return

    for process in processes:
        result = exec_utils.run_command(
            ["taskkill.exe", "/IM", process, "/F", "/T"],
            event="terminate_process",
            timeout=timeout,
            human_message=f"Terminating {process}",
            extra={"process_name": process},
        )

        if result.returncode == 127:
            human_logger.debug(
                "taskkill.exe is unavailable; skipping termination for %s", process
            )
            continue

        if result.timed_out:
            human_logger.warning("Timed out attempting to stop %s", process)
            machine_logger.warning(
                "terminate_process_timeout",
                extra={
                    "event": "terminate_process_timeout",
                    "process_name": process,
                    "stdout": result.stdout,
                    "stderr": result.stderr,
                },
            )
            continue

        if result.returncode == 0 and not result.error:
            human_logger.info("Terminated %s", process)
        else:
            human_logger.debug(
                "taskkill exited with %s for %s: %s",
                result.returncode,
                process,
                result.stderr.strip(),
            )


def terminate_process_patterns(patterns: Sequence[str], *, timeout: int = 30) -> None:
    """!
    @brief Terminate processes that match the provided wildcard patterns.
    @details Uses :func:`enumerate_processes` to expand patterns and forwards
    the resulting process list to :func:`terminate_office_processes`.
    @param patterns Wildcard expressions identifying Office executables.
    @param timeout Maximum seconds for enumeration and termination commands.
    """

    matches = enumerate_processes(patterns, timeout=timeout)
    if matches:
        terminate_office_processes(matches, timeout=timeout)
