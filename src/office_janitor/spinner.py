"""!
@brief Global persistent spinner for showing current task status.
@details Provides a dedicated status line at the bottom of the console that
never interferes with log output. Uses ANSI escape sequences to maintain
a separate scrolling region for logs above the status line.

Also provides SIGINT (Ctrl+C) handling for instant termination, including
killing any tracked subprocesses.
"""

from __future__ import annotations

import atexit
import os
import signal
import subprocess
import sys
import threading
import time
from typing import Any, Callable

# ---------------------------------------------------------------------------
# Global cancellation flag - checked by all blocking operations
# ---------------------------------------------------------------------------

_cancelled = threading.Event()


def is_cancelled() -> bool:
    """Check if cancellation has been requested (Ctrl+C pressed)."""
    return _cancelled.is_set()


def request_cancellation() -> None:
    """Request cancellation of all operations."""
    _cancelled.set()


def check_cancelled() -> None:
    """Raise KeyboardInterrupt if cancellation was requested."""
    if _cancelled.is_set():
        raise KeyboardInterrupt("Operation cancelled")


# Spinner animation frames - use ASCII-safe characters for Windows compatibility
_SPINNER_FRAMES = ("|", "/", "-", "\\")

# Global spinner state
_spinner_lock = threading.Lock()
_current_task: str | None = None
_task_start_time: float = 0.0
_spinner_idx: int = 0
_spinner_enabled: bool = True
_spinner_thread: threading.Thread | None = None
_spinner_stop_event: threading.Event | None = None
_status_line_active: bool = False  # Whether we've reserved a line for status
_output_paused: bool = False  # Whether output is paused (status line cleared for logging)

# Track active subprocesses for cleanup on SIGINT
_active_processes: set[subprocess.Popen[Any]] = set()
_process_lock = threading.Lock()

# Original signal handler (to restore if needed)
_original_sigint_handler: signal.Handlers | None = None
_sigint_installed: bool = False


def _format_elapsed(seconds: float) -> str:
    """Format elapsed seconds into a human-readable string."""
    if seconds < 60:
        return f"{int(seconds)}s"
    elif seconds < 3600:
        mins = int(seconds // 60)
        secs = int(seconds % 60)
        return f"{mins}m {secs}s" if secs else f"{mins}m"
    else:
        hours = int(seconds // 3600)
        mins = int((seconds % 3600) // 60)
        return f"{hours}h {mins}m" if mins else f"{hours}h"


def _get_terminal_width() -> int:
    """Get terminal width, defaulting to 80 if unknown."""
    try:
        import shutil

        return shutil.get_terminal_size().columns
    except Exception:
        return 80


def _get_terminal_height() -> int:
    """Get terminal height, defaulting to 24 if unknown."""
    try:
        import shutil

        return shutil.get_terminal_size().lines
    except Exception:
        return 24


def _draw_status_line() -> None:
    """Draw the status line - always visible at the current position."""
    global _status_line_active

    if not _spinner_enabled or _current_task is None:
        return

    frame = _SPINNER_FRAMES[_spinner_idx % len(_SPINNER_FRAMES)]

    elapsed = time.monotonic() - _task_start_time
    elapsed_str = _format_elapsed(elapsed)

    status = f"{frame} {_current_task}... ({elapsed_str})"

    # Truncate if too long for terminal
    width = _get_terminal_width()
    if len(status) > width - 1:
        status = status[: width - 4] + "..."

    # Always write on current line with carriage return
    # Use cyan color for visibility
    sys.stdout.write(f"\r\x1b[2K\x1b[36m{status}\x1b[0m")
    sys.stdout.flush()
    _status_line_active = True


def _clear_status_line() -> None:
    """Clear the status line content and move to next line."""
    global _status_line_active

    if _status_line_active:
        # Clear the spinner line and move to next line for clean output
        sys.stdout.write("\r\x1b[2K")
        sys.stdout.flush()
        _status_line_active = False


def _finalize_status_line() -> None:
    """Finalize the status line - print newline so it stays visible."""
    global _status_line_active

    if _status_line_active:
        # Just print newline to preserve the last status
        sys.stdout.write("\n")
        sys.stdout.flush()
        _status_line_active = False


def _spinner_loop() -> None:
    """Background thread loop that updates the spinner."""
    global _spinner_idx

    while _spinner_stop_event is not None and not _spinner_stop_event.is_set():
        with _spinner_lock:
            if _current_task is not None and _spinner_enabled and not _output_paused:
                _spinner_idx += 1
                _draw_status_line()
        _spinner_stop_event.wait(0.1)  # Update every 100ms


# ---------------------------------------------------------------------------
# SIGINT / Ctrl+C handling
# ---------------------------------------------------------------------------


def _sigint_handler(signum: int, frame: object) -> None:
    """
    Handle SIGINT (Ctrl+C) for instant termination.

    Sets cancellation flag, clears spinner, kills subprocesses, and exits.
    """
    # Set the cancellation flag FIRST - this unblocks waiting operations
    request_cancellation()

    # Clear status line
    with _spinner_lock:
        _clear_status_line()

    # Print cancellation message
    sys.stdout.write("\n\033[33m[INTERRUPTED]\033[0m Ctrl+C received, terminating...\n")
    sys.stdout.flush()

    # Kill all tracked subprocesses immediately (no waiting)
    with _process_lock:
        for proc in list(_active_processes):
            try:
                if proc.poll() is None:  # Still running
                    proc.kill()  # Immediate kill, no graceful terminate
            except Exception:
                pass  # Best effort cleanup
        _active_processes.clear()

    # Signal spinner thread to stop
    if _spinner_stop_event is not None:
        _spinner_stop_event.set()

    # Exit immediately - use os._exit to bypass any blocking cleanup
    os._exit(130)  # 128 + SIGINT(2)


def install_sigint_handler() -> None:
    """
    Install the SIGINT handler for instant Ctrl+C termination.

    Call this once at application startup to enable instant termination.
    Safe to call multiple times (only installs once).
    """
    global _original_sigint_handler, _sigint_installed

    if _sigint_installed:
        return

    _original_sigint_handler = signal.signal(signal.SIGINT, _sigint_handler)
    _sigint_installed = True


def uninstall_sigint_handler() -> None:
    """Restore the original SIGINT handler."""
    global _original_sigint_handler, _sigint_installed

    if not _sigint_installed:
        return

    if _original_sigint_handler is not None:
        signal.signal(signal.SIGINT, _original_sigint_handler)
    _original_sigint_handler = None
    _sigint_installed = False


def register_process(proc: subprocess.Popen[Any]) -> None:
    """
    Register a subprocess for cleanup on SIGINT.

    @param proc The Popen object to track.
    """
    with _process_lock:
        _active_processes.add(proc)


def unregister_process(proc: subprocess.Popen[Any]) -> None:
    """
    Unregister a subprocess (call when it completes normally).

    @param proc The Popen object to stop tracking.
    """
    with _process_lock:
        _active_processes.discard(proc)


def kill_all_processes() -> int:
    """
    Kill all tracked subprocesses.

    @return Number of processes that were killed.
    """
    killed = 0
    with _process_lock:
        for proc in list(_active_processes):
            try:
                if proc.poll() is None:  # Still running
                    proc.terminate()
                    try:
                        proc.wait(timeout=2.0)
                    except subprocess.TimeoutExpired:
                        proc.kill()
                    killed += 1
            except Exception:
                pass
        _active_processes.clear()
    return killed


# ---------------------------------------------------------------------------
# Spinner thread management
# ---------------------------------------------------------------------------


def start_spinner_thread() -> None:
    """Start the background spinner animation thread."""
    global _spinner_thread, _spinner_stop_event

    if _spinner_thread is not None and _spinner_thread.is_alive():
        return  # Already running

    # Also install SIGINT handler when spinner starts
    install_sigint_handler()

    _spinner_stop_event = threading.Event()
    _spinner_thread = threading.Thread(target=_spinner_loop, daemon=True, name="spinner")
    _spinner_thread.start()


def stop_spinner_thread() -> None:
    """Stop the background spinner animation thread."""
    global _spinner_thread, _spinner_stop_event

    if _spinner_stop_event is not None:
        _spinner_stop_event.set()

    if _spinner_thread is not None:
        _spinner_thread.join(timeout=0.5)
        _spinner_thread = None
        _spinner_stop_event = None

    # Finalize the status line (preserve it, don't clear)
    with _spinner_lock:
        _finalize_status_line()


def set_task(task_name: str | None) -> None:
    """
    Set the current task being worked on.

    @param task_name The task name to display, or None to clear.
    """
    global _current_task, _task_start_time, _spinner_idx

    with _spinner_lock:
        if task_name is None:
            _clear_status_line()
        _current_task = task_name
        if task_name is not None:
            _task_start_time = time.monotonic()
            _spinner_idx = 0
            # Draw immediately so status appears right away
            _draw_status_line()


def clear_task() -> None:
    """Clear the current task (equivalent to set_task(None))."""
    set_task(None)


def pause_for_output() -> None:
    """
    Prepare for log output by clearing current line and moving to new line.
    Call resume_after_output() when done printing.
    """
    global _output_paused
    with _spinner_lock:
        if not _output_paused:
            if _status_line_active:
                # Clear spinner line, print newline so logs go above
                sys.stdout.write("\r\x1b[2K")
                sys.stdout.flush()
            _output_paused = True


def resume_after_output() -> None:
    """
    Redraw the status line after log output is complete.
    """
    global _output_paused
    with _spinner_lock:
        if _output_paused:
            _output_paused = False
            if _current_task is not None and _spinner_enabled:
                _draw_status_line()


def spinner_print(message: str, **kwargs: object) -> None:
    """
    Print a message, automatically handling the status line.
    Use this instead of print() when the spinner may be active.
    """
    pause_for_output()
    try:
        print(message, **kwargs)
    finally:
        resume_after_output()


def enable_spinner(enabled: bool = True) -> None:
    """Enable or disable the spinner globally."""
    global _spinner_enabled
    with _spinner_lock:
        if not enabled:
            _clear_status_line()
        _spinner_enabled = enabled


def is_spinner_enabled() -> bool:
    """Check if spinner is currently enabled."""
    return _spinner_enabled


def get_current_task() -> str | None:
    """Get the current task name, or None if no task is set."""
    return _current_task


class TrackedProcess:
    """
    Context manager for tracking a subprocess for SIGINT cleanup.

    Usage:
        with TrackedProcess(proc):
            stdout, stderr = proc.communicate()
        # Process is automatically unregistered when done
    """

    def __init__(self, proc: subprocess.Popen[Any]) -> None:
        self._proc = proc

    def __enter__(self) -> subprocess.Popen[Any]:
        register_process(self._proc)
        return self._proc

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_val: BaseException | None,
        exc_tb: object,
    ) -> None:
        unregister_process(self._proc)
        return None


class SpinnerTask:
    """
    Context manager for setting a spinner task.

    Usage:
        with SpinnerTask("Detecting installations"):
            # do work
        # spinner automatically clears when exiting
    """

    def __init__(self, task_name: str) -> None:
        self._task_name = task_name
        self._previous_task: str | None = None

    def __enter__(self) -> "SpinnerTask":
        self._previous_task = get_current_task()
        set_task(self._task_name)
        return self

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc_val: BaseException | None,
        exc_tb: object,
    ) -> None:
        # Restore previous task (or clear if there wasn't one)
        set_task(self._previous_task)
        return None


# Register cleanup on exit
def _cleanup_spinner() -> None:
    """Cleanup spinner on program exit."""
    with _spinner_lock:
        _clear_status_line()
    stop_spinner_thread()


atexit.register(_cleanup_spinner)


# ---------------------------------------------------------------------------
# Interruptible waiting utilities
# ---------------------------------------------------------------------------


def wait_interruptible(timeout: float = 0.1) -> bool:
    """
    Sleep for up to `timeout` seconds, but return early if cancelled.

    @param timeout Maximum time to wait in seconds.
    @return True if cancelled, False if timeout expired normally.
    """
    return _cancelled.wait(timeout)


def wait_for_future(
    future: Any,  # concurrent.futures.Future
    timeout: float | None = None,
    poll_interval: float = 0.1,
) -> Any:
    """
    Wait for a future to complete, checking for cancellation between polls.

    @param future The Future object to wait for.
    @param timeout Maximum total time to wait (None = forever).
    @param poll_interval How often to check for cancellation.
    @return The future's result.
    @raises KeyboardInterrupt if cancellation is requested.
    @raises TimeoutError if timeout expires before completion.
    """
    import concurrent.futures

    start = time.monotonic()
    while True:
        # Check for cancellation
        if is_cancelled():
            future.cancel()
            raise KeyboardInterrupt("Operation cancelled")

        try:
            return future.result(timeout=poll_interval)
        except concurrent.futures.TimeoutError:
            # Check overall timeout
            if timeout is not None:
                elapsed = time.monotonic() - start
                if elapsed >= timeout:
                    future.cancel()
                    raise TimeoutError(f"Timed out after {elapsed:.1f}s")
            # Continue polling
            continue


def wait_for_futures(
    futures: dict[str, Any],  # dict of name -> Future
    timeout: float | None = None,
    poll_interval: float = 0.1,
) -> dict[str, Any]:
    """
    Wait for multiple futures, checking for cancellation between polls.

    @param futures Dict mapping names to Future objects.
    @param timeout Maximum total time to wait (None = forever).
    @param poll_interval How often to check for cancellation.
    @return Dict mapping names to results.
    @raises KeyboardInterrupt if cancellation is requested.
    @raises Exception if any future raised an exception.
    """
    import concurrent.futures

    results: dict[str, Any] = {}
    errors: dict[str, Exception] = {}
    pending = set(futures.keys())
    start = time.monotonic()

    while pending:
        # Check for cancellation
        if is_cancelled():
            for name in pending:
                futures[name].cancel()
            raise KeyboardInterrupt("Operation cancelled")

        # Check each pending future
        for name in list(pending):
            future = futures[name]
            if future.done():
                try:
                    results[name] = future.result()
                except Exception as e:
                    errors[name] = e
                pending.remove(name)

        if pending:
            # Check overall timeout
            if timeout is not None:
                elapsed = time.monotonic() - start
                if elapsed >= timeout:
                    for name in pending:
                        futures[name].cancel()
                    raise TimeoutError(f"Timed out after {elapsed:.1f}s")

            # Sleep briefly before next poll
            time.sleep(poll_interval)

    # Re-raise any errors that occurred
    if errors:
        # Raise the first error encountered
        first_name, first_error = next(iter(errors.items()))
        raise first_error

    return results
