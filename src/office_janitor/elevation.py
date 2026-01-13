"""!
@brief Elevation and user-context helpers.
@details Provides detection and relaunch utilities to request administrative
rights when required, as well as helpers to execute commands under a limited
user context for parity with VBS flows that relaunch as the interactive user.
All helpers rely on the standard library and log through :mod:`exec_utils`.
"""

from __future__ import annotations

import ctypes
import os
import shutil
import subprocess
import sys
from collections.abc import Iterable, Mapping, Sequence

from . import exec_utils, logging_ext


def is_admin() -> bool:
    """!
    @brief Determine whether the current process token has administrative rights.
    """

    if os.name != "nt":
        return False
    try:
        shell32 = ctypes.windll.shell32
        return bool(shell32.IsUserAnAdmin())
    except Exception:
        return False


def current_username() -> str:
    """!
    @brief Return the current user name best-effort.
    """

    for candidate in (
        os.getlogin,
        lambda: os.environ.get("USERNAME"),
        lambda: os.environ.get("USER"),
    ):
        try:
            value = candidate()
        except Exception:
            value = None
        if value:
            return str(value)
    return ""


# Environment variable used to signal the process was auto-elevated
_ELEVATED_MARKER_ENV = "OFFICE_JANITOR_ELEVATED"


def was_auto_elevated() -> bool:
    """!
    @brief Check if this process was launched via auto-elevation.
    @details Returns True if the process was relaunched with elevated privileges
    by relaunch_as_admin(), indicating the console window should be kept open.
    """
    return os.environ.get(_ELEVATED_MARKER_ENV, "") == "1"


def relaunch_as_admin(argv: Sequence[str] | None = None) -> bool:
    """!
    @brief Relaunch the current interpreter with administrative rights.
    @details Sets an environment marker so the elevated process knows it was
    auto-elevated and should keep the console window open on exit.
    @returns ``True`` when the relaunch request was issued successfully.
    """

    if os.name != "nt":
        return False
    try:
        shell32 = ctypes.windll.shell32
    except Exception:
        return False

    arguments = list(argv) if argv is not None else list(sys.argv[1:])
    # Prepend environment marker to the arguments
    # This passes the marker to the elevated process
    marker_arg = "--_elevated-marker"
    arguments = [marker_arg] + arguments
    params = subprocess.list2cmdline(arguments)
    result = shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
    return int(result) > 32


def pause_if_elevated(exit_code: int = 0) -> None:
    """!
    @brief Pause the console if this was an auto-elevated process.
    @details When the process was launched via UAC elevation, the console window
    would normally close immediately on exit. This function pauses execution to
    allow the user to see the output before closing.
    @param exit_code The exit code that will be returned (shown in message).
    """
    if was_auto_elevated():
        print()
        if exit_code == 0:
            print("=" * 60)
            print("Operation completed successfully.")
            print("=" * 60)
        else:
            print("=" * 60)
            print(f"Operation finished with exit code: {exit_code}")
            print("=" * 60)
        print()
        try:
            input("Press Enter to close this window...")
        except (EOFError, KeyboardInterrupt):
            pass  # Handle non-interactive or Ctrl+C gracefully


def run_as_limited_user(
    command: Sequence[str] | str,
    *,
    event: str = "runas_limited",
    human_message: str | None = None,
    dry_run: bool = False,
    env: Mapping[str, str] | None = None,
    inherit_env: bool = True,
    env_overrides: Mapping[str, str] | None = None,
    env_remove: Iterable[str] | None = None,
    cwd: str | None = None,
) -> exec_utils.CommandResult:
    """!
    @brief Attempt to execute ``command`` under a limited user context.
    @details Uses ``runas.exe /trustlevel:0x20000`` when available on Windows.
    If the helper is unavailable, falls back to normal execution while logging a
    warning so callers know de-elevation was not applied.
    """

    human_logger = logging_ext.get_human_logger()
    runas_path = shutil.which("runas.exe") if os.name == "nt" else None

    if runas_path:
        runas_cmd: list[str] = [
            runas_path,
            "/trustlevel:0x20000",
        ]
        if isinstance(command, str):
            runas_cmd.append(command)
        else:
            runas_cmd.append(" ".join(str(part) for part in command))
        human_logger.info("Executing command as limited user via runas.exe")
        return exec_utils.run_command(
            runas_cmd,
            event=event,
            dry_run=dry_run,
            human_message=human_message or "Running command as limited user",
            env=env,
            inherit_env=inherit_env,
            env_overrides=env_overrides,
            env_remove=env_remove,
            cwd=cwd,
        )

    human_logger.warning("runas.exe not available; executing command without de-elevation.")
    return exec_utils.run_command(
        command,
        event=event,
        dry_run=dry_run,
        human_message=human_message or "Running command without de-elevation (fallback)",
        env=env,
        inherit_env=inherit_env,
        env_overrides=env_overrides,
        env_remove=env_remove,
        cwd=cwd,
    )


__all__ = [
    "is_admin",
    "current_username",
    "relaunch_as_admin",
    "run_as_limited_user",
    "was_auto_elevated",
    "pause_if_elevated",
]
