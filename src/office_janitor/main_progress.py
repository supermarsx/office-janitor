"""!
@file main_progress.py
@brief Progress logging utilities for the Office Janitor CLI.
@details Provides Linux init-style progress output with timestamps,
thread-safe printing, and console setup utilities.
"""

from __future__ import annotations

import ctypes
import os
import threading
import time
from typing import TYPE_CHECKING

from . import spinner

if TYPE_CHECKING:
    pass

__all__ = [
    "get_main_start_time",
    "set_main_start_time",
    "get_elapsed_secs",
    "progress",
    "progress_ok",
    "progress_fail",
    "progress_skip",
    "enable_vt_mode_if_possible",
]

# Module-level state for consistent timestamps across all progress output
_MAIN_START_TIME: float = time.perf_counter()
_PROGRESS_LOCK = threading.Lock()
_PENDING_LINE_OWNER: int | None = None


def get_main_start_time() -> float:
    """!
    @brief Get the current main start time.
    @returns The timestamp when main() was called.
    """
    return _MAIN_START_TIME


def set_main_start_time(start_time: float) -> None:
    """!
    @brief Set the main start time for progress tracking.
    @param start_time The timestamp to use as the start time.
    """
    global _MAIN_START_TIME
    _MAIN_START_TIME = start_time


def get_elapsed_secs() -> float:
    """!
    @brief Get elapsed seconds since program start.
    @returns Elapsed time in seconds with high precision.
    """
    return time.perf_counter() - _MAIN_START_TIME


def progress(
    message: str,
    *,
    indent: int = 0,
    newline: bool = True,
) -> None:
    """!
    @brief Print a timestamped progress message in Linux init style.
    @param message The message to print.
    @param indent Indentation level (each level adds 2 spaces).
    @param newline Whether to end with a newline (False for status pending).
    """
    global _PENDING_LINE_OWNER
    prefix = "  " * indent
    timestamp = f"[{get_elapsed_secs():12.6f}]"
    text = f"{timestamp} {prefix}{message}"

    with _PROGRESS_LOCK:
        # Clear incomplete line flag before printing
        spinner.clear_incomplete_line()
        spinner.pause_for_output()
        try:
            current_thread = threading.get_ident()
            # Complete any previous incomplete line from a different thread
            if _PENDING_LINE_OWNER is not None and _PENDING_LINE_OWNER != current_thread:
                print(flush=True)

            if newline:
                print(text, flush=True)
                _PENDING_LINE_OWNER = None
            else:
                print(text, end="", flush=True)
                _PENDING_LINE_OWNER = current_thread
        finally:
            spinner.resume_after_output()


def progress_ok(detail: str | None = None) -> None:
    """!
    @brief Print OK status in Linux init style [  OK  ].
    @param detail Optional detail to append after the status.
    """
    global _PENDING_LINE_OWNER
    with _PROGRESS_LOCK:
        # Clear incomplete line flag before printing
        spinner.clear_incomplete_line()
        spinner.pause_for_output()
        try:
            current_thread = threading.get_ident()
            if _PENDING_LINE_OWNER == current_thread:
                suffix = f" ({detail})" if detail else ""
                print(f" [  \033[32mOK\033[0m  ]{suffix}", flush=True)
                _PENDING_LINE_OWNER = None
            else:
                if _PENDING_LINE_OWNER is not None:
                    print(flush=True)
                suffix = f" ({detail})" if detail else ""
                print(
                    f"[{get_elapsed_secs():12.6f}]  [  \033[32mOK\033[0m  ]{suffix}",
                    flush=True,
                )
                _PENDING_LINE_OWNER = None
        finally:
            spinner.resume_after_output()


def progress_fail(reason: str | None = None) -> None:
    """!
    @brief Print FAILED status in Linux init style [FAILED].
    @param reason Optional reason to append after the status.
    """
    global _PENDING_LINE_OWNER
    with _PROGRESS_LOCK:
        # Clear incomplete line flag before printing
        spinner.clear_incomplete_line()
        spinner.pause_for_output()
        try:
            current_thread = threading.get_ident()
            if _PENDING_LINE_OWNER == current_thread:
                suffix = f" ({reason})" if reason else ""
                print(f" [\033[31mFAILED\033[0m]{suffix}", flush=True)
                _PENDING_LINE_OWNER = None
            else:
                if _PENDING_LINE_OWNER is not None:
                    print(flush=True)
                suffix = f" ({reason})" if reason else ""
                print(
                    f"[{get_elapsed_secs():12.6f}]  [\033[31mFAILED\033[0m]{suffix}",
                    flush=True,
                )
                _PENDING_LINE_OWNER = None
        finally:
            spinner.resume_after_output()


def progress_skip(reason: str | None = None) -> None:
    """!
    @brief Print SKIP status in Linux init style [ SKIP ].
    @param reason Optional reason to append after the status.
    """
    global _PENDING_LINE_OWNER
    with _PROGRESS_LOCK:
        # Clear incomplete line flag before printing
        spinner.clear_incomplete_line()
        spinner.pause_for_output()
        try:
            current_thread = threading.get_ident()
            if _PENDING_LINE_OWNER == current_thread:
                suffix = f" ({reason})" if reason else ""
                print(f" [ \033[33mSKIP\033[0m ]{suffix}", flush=True)
                _PENDING_LINE_OWNER = None
            else:
                if _PENDING_LINE_OWNER is not None:
                    print(flush=True)
                suffix = f" ({reason})" if reason else ""
                print(
                    f"[{get_elapsed_secs():12.6f}]  [ \033[33mSKIP\033[0m ]{suffix}",
                    flush=True,
                )
                _PENDING_LINE_OWNER = None
        finally:
            spinner.resume_after_output()


def enable_vt_mode_if_possible() -> None:
    """!
    @brief Attempt to enable ANSI/VT processing on Windows consoles.
    @details Per the specification, the application should try to enable virtual
    terminal support so both the plain CLI and future TUI renderer can emit
    colorized output. Failures are silently ignored because the feature is
    optional.
    """
    if os.name != "nt":  # pragma: no cover - Windows behaviour only
        return

    try:
        from ctypes import wintypes

        kernel32 = ctypes.windll.kernel32
    except Exception:  # pragma: no cover - import/attribute errors on non-Windows
        return

    for std_handle in (-11, -12):  # STD_OUTPUT_HANDLE, STD_ERROR_HANDLE
        handle = kernel32.GetStdHandle(std_handle)
        if not handle:
            continue
        mode = wintypes.DWORD()
        if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
            kernel32.SetConsoleMode(handle, mode.value | 0x0004)


def get_progress_lock() -> threading.Lock:
    """!
    @brief Get the progress lock for thread-safe printing.
    @returns The module-level progress lock.
    """
    return _PROGRESS_LOCK
