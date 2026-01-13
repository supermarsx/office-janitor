"""!
@brief Process management helpers for Office scrubbing.
@details Provides helpers to discover and terminate running Office
processes. The module mirrors the behaviour from the legacy OffScrub
scripts with structured logging, user prompting, and timeout-aware
subprocess execution so automated runs stay safe.
"""

from __future__ import annotations

import fnmatch
from collections.abc import Iterable, Sequence
from typing import Callable

from . import exec_utils, logging_ext


def enumerate_processes(patterns: Iterable[str], *, timeout: int = 30) -> list[str]:
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

    expanded_patterns: list[str] = [
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
        human_logger.debug("tasklist exited with %s; skipping enumeration", listing.returncode)
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

    running: list[str] = []
    for line in listing.stdout.splitlines():
        stripped = line.strip()
        if not stripped or stripped.lower().startswith("image name"):
            continue
        name = stripped.split()[0].lower()
        if name not in running:
            running.append(name)

    matched: list[str] = []
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

    normalized = {name.lower() for name in remaining}
    if "outlook.exe" in normalized:
        reassurance = (
            "Outlook data providers are only suspended; OST/PST mail stores remain intact."
        )
        human_logger.warning(reassurance)
        logging_ext.emit_ui_event(
            "processes.outlook_reassurance",
            reassurance,
            status="Outlook data providers suspended; OST/PST files remain intact.",
        )

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

    processes: list[str] = [str(name).strip() for name in names if str(name).strip()]
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
            human_logger.debug("taskkill.exe is unavailable; skipping termination for %s", process)
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


def is_explorer_running(*, timeout: int = 10) -> bool:
    """!
    @brief Check if explorer.exe is currently running.
    @param timeout Maximum seconds for process query.
    @returns True if explorer.exe is running, False otherwise.
    """
    from . import exec_utils

    result = exec_utils.run_command(
        ["tasklist", "/FI", "IMAGENAME eq explorer.exe", "/FO", "CSV"],
        event="check_explorer",
        timeout=timeout,
    )

    if result.returncode != 0:
        return False

    output = (result.stdout or "").lower()
    return "explorer.exe" in output


def restart_explorer_if_needed(*, timeout: int = 10) -> bool:
    """!
    @brief Restart explorer.exe if it's not running.
    @details VBS equivalent: RestoreExplorer in OffScrubC2R.vbs.
    Called after shell integration cleanup that may have terminated
    explorer to release file locks.
    @param timeout Maximum seconds for commands.
    @returns True if explorer was restarted, False if already running.
    """
    from . import exec_utils, logging_ext

    human_logger = logging_ext.get_human_logger()

    if is_explorer_running(timeout=timeout):
        human_logger.debug("explorer.exe is already running")
        return False

    human_logger.info("Restarting explorer.exe")

    result = exec_utils.run_command(
        ["cmd.exe", "/c", "start", "explorer.exe"],
        event="restart_explorer",
        timeout=timeout,
    )

    if result.returncode == 0:
        human_logger.info("explorer.exe restarted successfully")
        return True

    human_logger.warning("Failed to restart explorer.exe: %d", result.returncode)
    return False


def terminate_all_office_processes(
    *,
    include_infrastructure: bool = True,
    dry_run: bool = False,
    timeout: int = 30,
) -> list[str]:
    """!
    @brief Terminate all known Office processes.
    @details Combines standard Office apps with infrastructure processes.
    VBS equivalent: CloseOfficeApps in OffScrub scripts.
    @param include_infrastructure If True, include C2R and MSI infrastructure processes.
    @param dry_run If True, only report what would be terminated.
    @param timeout Maximum seconds for termination commands.
    @returns List of process names that were terminated (or would be in dry-run).
    """
    from . import constants, logging_ext

    human_logger = logging_ext.get_human_logger()

    # Build the process list
    if include_infrastructure:
        processes = constants.ALL_OFFICE_PROCESSES
    else:
        processes = constants.DEFAULT_OFFICE_PROCESSES

    # Find which processes are actually running
    running = enumerate_processes(list(processes), timeout=timeout)

    if not running:
        human_logger.debug("No Office processes found running")
        return []

    human_logger.info("Found %d Office process(es) to terminate", len(running))

    if dry_run:
        for proc in running:
            human_logger.info("[DRY-RUN] Would terminate: %s", proc)
        return list(running)

    terminate_office_processes(running, timeout=timeout)
    return list(running)
