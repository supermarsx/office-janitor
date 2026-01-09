"""!
@brief Subprocess execution helpers with sanitised environments.
@details Centralises invocation of :func:`subprocess.run` so callers inherit
consistent logging, dry-run behaviour, and environment handling. The helper is
used across the project for PowerShell, command-line, and system utilities to
ensure telemetry remains uniform while also protecting subprocesses from leaked
virtual environment variables.
"""

from __future__ import annotations

import os
import subprocess
import time
from collections.abc import Iterable, Mapping, MutableMapping, Sequence
from dataclasses import dataclass
from typing import Any

from . import logging_ext

_SANITIZE_BLOCKLIST = {
    "PYTHONPATH",
    "PYTHONHOME",
    "PYTHONWARNINGS",
    "VIRTUAL_ENV",
    "PIP_REQUIRE_VIRTUALENV",
    "CONDA_PREFIX",
    "CONDA_DEFAULT_ENV",
    "PYENV_VERSION",
    "POETRY_ACTIVE",
    "__PYVENV_LAUNCHER__",
}


@dataclass
class CommandResult:
    """!
    @brief Outcome information from :func:`run_command`.
    @details Encapsulates the executed command, captured output streams,
    duration, and metadata describing dry-run or timeout states.
    """

    command: Sequence[str]
    returncode: int
    stdout: str
    stderr: str
    duration: float
    skipped: bool = False
    timed_out: bool = False
    error: str | None = None


_GLOBAL_TIMEOUT: float | None = None


def set_global_timeout(timeout_seconds: float | int | None) -> None:
    """!
    @brief Apply a global timeout cap for all subprocess calls.
    @details When set, :func:`run_command` will use the minimum of the caller
    supplied timeout and this global limit so the CLI ``--timeout`` flag can
    enforce per-step ceilings.
    """

    global _GLOBAL_TIMEOUT
    if timeout_seconds is None:
        _GLOBAL_TIMEOUT = None
        return
    try:
        parsed = float(timeout_seconds)
    except (TypeError, ValueError):
        _GLOBAL_TIMEOUT = None
    else:
        _GLOBAL_TIMEOUT = parsed if parsed > 0 else None


def _resolve_timeout(requested: float | int | None) -> float | int | None:
    """!
    @brief Determine the effective timeout considering the global override.
    """

    if _GLOBAL_TIMEOUT is None:
        return requested
    if requested is None:
        return _GLOBAL_TIMEOUT
    return min(_GLOBAL_TIMEOUT, requested)


def _build_call_payload(
    command_list: Sequence[str],
    *,
    timeout: float | int | None,
    cwd: str | None,
    extra: Mapping[str, object] | None,
) -> MutableMapping[str, object]:
    payload: MutableMapping[str, object] = {
        "command": list(command_list),
        "timeout": timeout,
    }
    if cwd:
        payload["cwd"] = cwd
    if extra:
        for key, value in extra.items():
            if key not in {"event", "result"}:
                payload[key] = value
    return payload


def _build_result_payload(
    *,
    return_code: int,
    duration: float,
    stdout: str,
    stderr: str,
    error: str | None = None,
    timed_out: bool = False,
) -> dict[str, object]:
    return {
        "rc": return_code,
        "duration_ms": round(duration * 1000, 3),
        "stdout": stdout,
        "stderr": stderr,
        "error": error,
        "timed_out": timed_out,
    }


def sanitize_environment(
    *,
    base_env: Mapping[str, str] | None = None,
    inherit: bool = True,
    extra: Mapping[str, str] | None = None,
    remove: Iterable[str] | None = None,
) -> MutableMapping[str, str]:
    """!
    @brief Produce a subprocess environment stripped of virtualenv artefacts.
    @details ``base_env`` defaults to :data:`os.environ` when ``inherit`` is
    ``True``. Sanitisation removes variables that commonly interfere with child
    processes, especially when the application is packaged with PyInstaller.
    Additional variables can be overlaid via ``extra`` or removed via
    ``remove``.
    @param base_env Source mapping to copy prior to sanitisation.
    @param inherit When ``True`` and ``base_env`` is ``None`` the host
    environment is used as a starting point.
    @param extra Mapping of overrides/augmentations applied after sanitisation.
    @param remove Additional variable names to drop after the default blocklist.
    @returns Mutable mapping ready for subprocess invocation.
    """

    if base_env is not None:
        environment: MutableMapping[str, str] = {
            str(k): str(v) for k, v in base_env.items() if v is not None
        }
    elif inherit:
        environment = {str(k): str(v) for k, v in os.environ.items() if v is not None}
    else:
        environment = {}

    for key in _SANITIZE_BLOCKLIST:
        environment.pop(key, None)

    if remove is not None:
        for key in remove:
            environment.pop(key, None)

    if extra:
        for key, value in extra.items():
            environment[str(key)] = str(value)

    return environment


def run_command(
    command: Sequence[str] | str,
    *,
    event: str,
    timeout: int | float | None = None,
    dry_run: bool = False,
    human_message: str | None = None,
    extra: Mapping[str, object] | None = None,
    env: Mapping[str, str] | None = None,
    inherit_env: bool = True,
    env_overrides: Mapping[str, str] | None = None,
    env_remove: Iterable[str] | None = None,
    cwd: str | None = None,
    check: bool = False,
) -> CommandResult:
    """!
    @brief Execute ``command`` with consistent logging and environment hygiene.
    @details Emits ``*_plan`` and ``*_result`` machine-log events, logs human
    friendly messaging, and supports dry-run mode which echoes the intended
    command without executing it. The child environment is sanitized to remove
    Python virtual environment artefacts unless explicitly overridden.
    @param command Sequence of command arguments.
    @param event Base name for structured log events.
    @param timeout Optional timeout (seconds) passed to :func:`subprocess.run`.
    @param dry_run When ``True`` no subprocess is spawned and the result is
    marked as ``skipped``.
    @param human_message Optional message emitted to the human logger before
    execution.
    @param extra Additional metadata merged into machine log payloads.
    @param env Explicit environment mapping to start from prior to sanitisation.
    @param inherit_env Whether to inherit :data:`os.environ` when ``env`` is
    ``None``.
    @param env_overrides Mapping applied after sanitisation.
    @param env_remove Additional variables to remove from the environment.
    @param cwd Working directory supplied to :func:`subprocess.run`.
    @param check When ``True`` non-zero exit codes raise
    :class:`subprocess.CalledProcessError` after logging the result.
    @returns :class:`CommandResult` describing the observed outcome.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    if isinstance(command, str):
        command_list = [command]
    else:
        command_list = [str(part) for part in command]

    effective_timeout: Any = _resolve_timeout(timeout)

    metadata: MutableMapping[str, object] = {
        "event": f"{event}_plan",
        "call": _build_call_payload(
            command_list,
            timeout=effective_timeout,
            cwd=cwd,
            extra=extra,
        ),
        "dry_run": dry_run,
    }
    machine_logger.info(f"{event}_plan", extra=dict(metadata))

    if dry_run:
        if human_message:
            human_logger.info("%s [dry-run]", human_message)
        else:
            human_logger.info("Dry-run: would execute %s", " ".join(command_list))
        dry_meta: MutableMapping[str, object] = {
            "event": f"{event}_dry_run",
            "call": _build_call_payload(
                command_list,
                timeout=effective_timeout,
                cwd=cwd,
                extra=extra,
            ),
            "result": _build_result_payload(
                return_code=0,
                duration=0.0,
                stdout="",
                stderr="",
                error=None,
                timed_out=False,
            ),
        }
        machine_logger.info(f"{event}_dry_run", extra=dict(dry_meta))
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

    sanitized_env = sanitize_environment(
        base_env=env,
        inherit=inherit_env,
        extra=env_overrides,
        remove=env_remove,
    )

    start = time.monotonic()
    try:
        completed = subprocess.run(  # noqa: S603 - intentional command execution
            command_list,
            capture_output=True,
            text=True,
            timeout=effective_timeout,
            check=False,
            env=sanitized_env,
            cwd=cwd,
        )
    except FileNotFoundError as exc:
        duration = time.monotonic() - start
        human_logger.error("Command not found: %s", command_list[0])
        failure_meta: MutableMapping[str, object] = {
            "event": f"{event}_missing",
            "call": _build_call_payload(
                command_list,
                timeout=effective_timeout,
                cwd=cwd,
                extra=extra,
            ),
            "result": _build_result_payload(
                return_code=127,
                duration=duration,
                stdout="",
                stderr="",
                error=str(exc),
                timed_out=False,
            ),
        }
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
            "call": _build_call_payload(
                command_list,
                timeout=effective_timeout,
                cwd=cwd,
                extra=extra,
            ),
            "result": _build_result_payload(
                return_code=exc.returncode if hasattr(exc, "returncode") else 1,
                duration=duration,
                stdout=str(exc.stdout or ""),
                stderr=str(exc.stderr or ""),
                error="timeout",
                timed_out=True,
            ),
        }
        machine_logger.error(f"{event}_timeout", extra=dict(failure_meta))
        return CommandResult(
            command=command_list,
            returncode=1,
            stdout=str(exc.stdout or ""),
            stderr=str(exc.stderr or ""),
            duration=duration,
            timed_out=True,
            error="timeout",
        )
    except OSError as exc:
        duration = time.monotonic() - start
        human_logger.error("Failed to execute %s: %s", command_list[0], exc)
        failure_meta = {
            "event": f"{event}_error",
            "call": _build_call_payload(
                command_list,
                timeout=effective_timeout,
                cwd=cwd,
                extra=extra,
            ),
            "result": _build_result_payload(
                return_code=1,
                duration=duration,
                stdout="",
                stderr="",
                error=str(exc),
                timed_out=False,
            ),
        }
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
        "call": _build_call_payload(
            command_list,
            timeout=effective_timeout,
            cwd=cwd,
            extra=extra,
        ),
        "result": _build_result_payload(
            return_code=completed.returncode,
            duration=duration,
            stdout=str(completed.stdout),
            stderr=str(completed.stderr),
            error=None,
            timed_out=False,
        ),
    }
    machine_logger.info(f"{event}_result", extra=dict(result_meta))

    if completed.returncode != 0:
        human_logger.warning("Command %s exited with %s", command_list[0], completed.returncode)

    result = CommandResult(
        command=command_list,
        returncode=completed.returncode,
        stdout=completed.stdout,
        stderr=completed.stderr,
        duration=duration,
    )

    if check and completed.returncode != 0:
        raise subprocess.CalledProcessError(
            completed.returncode,
            command_list,
            output=completed.stdout,
            stderr=completed.stderr,
        )

    return result
