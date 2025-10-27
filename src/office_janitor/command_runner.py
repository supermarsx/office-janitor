"""!
@brief Shared subprocess execution helpers.
@details Provides a consistent wrapper around :func:`subprocess.run` that
records structured telemetry for command invocations, including execution
plans, durations, and failure metadata. Callers in the uninstall pipeline use
this module so command logging stays uniform across MSI and Click-to-Run flows.
"""
from __future__ import annotations

import subprocess
import time
from dataclasses import dataclass
from typing import Mapping, MutableMapping, Sequence

from . import logging_ext


@dataclass
class CommandResult:
    """!
    @brief Outcome metadata returned by :func:`run_command`.
    @details Captures the executed command, return code, collected streams, and
    runtime characteristics. ``skipped`` is ``True`` when dry-run mode bypassed
    the subprocess execution. ``timed_out`` is ``True`` when the command exceeded
    the requested timeout.
    """

    command: Sequence[str]
    returncode: int
    stdout: str
    stderr: str
    duration: float
    skipped: bool = False
    timed_out: bool = False
    error: str | None = None


def run_command(
    command: Sequence[str],
    *,
    event: str,
    timeout: int | float | None = None,
    dry_run: bool = False,
    human_message: str | None = None,
    extra: Mapping[str, object] | None = None,
) -> CommandResult:
    """!
    @brief Execute ``command`` while emitting structured telemetry records.
    @details The helper logs a ``*_plan`` event prior to invocation and a
    ``*_result`` event once the process exits (or a ``*_timeout``/``*_missing``
    failure). Human-oriented messaging is routed through the human logger so the
    console output mirrors the machine telemetry.
    @param command Command sequence to execute.
    @param event Base event identifier recorded in machine logs.
    @param timeout Optional timeout in seconds for the subprocess.
    @param dry_run When ``True`` skip execution and return a ``skipped`` result.
    @param human_message Optional message logged to the human channel before
    execution.
    @param extra Mapping merged into machine log ``extra`` payloads.
    @returns :class:`CommandResult` describing the observed outcome.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    command_list = [str(part) for part in command]
    metadata: MutableMapping[str, object] = {"event": f"{event}_plan", "command": command_list}
    if extra:
        metadata.update(extra)
    machine_logger.info(f"{event}_plan", extra=dict(metadata))

    if dry_run:
        human_logger.info(
            human_message or "Dry-run: would execute %s", " ".join(command_list)
        )
        dry_metadata: MutableMapping[str, object] = {
            "event": f"{event}_dry_run",
            "command": command_list,
        }
        if extra:
            dry_metadata.update(extra)
        machine_logger.info(f"{event}_dry_run", extra=dict(dry_metadata))
        return CommandResult(
            command=command_list,
            returncode=0,
            stdout="",
            stderr="",
            duration=0.0,
            skipped=True,
        )

    if human_message:
        human_logger.info(human_message)

    start = time.monotonic()
    try:
        completed = subprocess.run(  # noqa: S603 - intentional command execution
            command_list,
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
    except FileNotFoundError as exc:
        duration = time.monotonic() - start
        human_logger.error("Command not found: %s", command_list[0])
        failure_meta: MutableMapping[str, object] = {
            "event": f"{event}_missing",
            "command": command_list,
            "duration": duration,
            "error": str(exc),
        }
        if extra:
            failure_meta.update(extra)
        machine_logger.error(f"{event}_missing", extra=dict(failure_meta))
        return CommandResult(
            command=command_list,
            returncode=127,
            stdout="",
            stderr="",
            duration=duration,
            error=str(exc),
        )
    except subprocess.TimeoutExpired as exc:
        duration = time.monotonic() - start
        human_logger.error("Command timed out after %.1fs: %s", duration, command_list[0])
        failure_meta = {
            "event": f"{event}_timeout",
            "command": command_list,
            "duration": duration,
            "stdout": exc.stdout or "",
            "stderr": exc.stderr or "",
        }
        if extra:
            failure_meta.update(extra)
        machine_logger.error(f"{event}_timeout", extra=dict(failure_meta))
        return CommandResult(
            command=command_list,
            returncode=1,
            stdout=exc.stdout or "",
            stderr=exc.stderr or "",
            duration=duration,
            timed_out=True,
            error="timeout",
        )
    except OSError as exc:
        duration = time.monotonic() - start
        human_logger.error("Failed to execute %s: %s", command_list[0], exc)
        failure_meta = {
            "event": f"{event}_error",
            "command": command_list,
            "duration": duration,
            "error": str(exc),
        }
        if extra:
            failure_meta.update(extra)
        machine_logger.error(f"{event}_error", extra=dict(failure_meta))
        return CommandResult(
            command=command_list,
            returncode=1,
            stdout="",
            stderr="",
            duration=duration,
            error=str(exc),
        )

    duration = time.monotonic() - start
    result_meta: MutableMapping[str, object] = {
        "event": f"{event}_result",
        "command": command_list,
        "return_code": completed.returncode,
        "stdout": completed.stdout,
        "stderr": completed.stderr,
        "duration": duration,
    }
    if extra:
        result_meta.update(extra)
    machine_logger.info(f"{event}_result", extra=dict(result_meta))

    if completed.returncode != 0:
        human_logger.warning(
            "Command %s exited with %s", command_list[0], completed.returncode
        )

    return CommandResult(
        command=command_list,
        returncode=completed.returncode,
        stdout=completed.stdout,
        stderr=completed.stderr,
        duration=duration,
    )
