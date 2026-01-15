"""!
@brief Global persistent spinner for showing current task status.
@details Provides a dedicated status line at the bottom of the console that
never interferes with log output. Uses ANSI escape sequences to maintain
a separate scrolling region for logs above the status line.

IMPORTANT: Spinner output is written DIRECTLY to the console (sys.__stdout__)
and NEVER goes through the logging system. This ensures spinner animations
never appear in log files.

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
from typing import Any, Callable, TextIO

# ---------------------------------------------------------------------------
# Direct console output - NEVER use logging for spinner
# ---------------------------------------------------------------------------


# Use the original stdout that bypasses any redirections
# This ensures spinner output NEVER goes to log files
def _get_console() -> TextIO:
    """Get the real console output stream, bypassing any redirections."""
    # sys.__stdout__ is the original stdout before any redirections
    # This ensures spinner output goes ONLY to the actual console
    return sys.__stdout__ if sys.__stdout__ is not None else sys.stdout


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
_pending_incomplete_line: bool = False  # Whether there's an incomplete line awaiting continuation

# Multi-task tracking for parallel operations
# Maps task_name -> start_time (allows multiple concurrent tasks)
_active_tasks: dict[str, float] = {}

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
    """Draw the status line - always visible at the current position.

    IMPORTANT: Output goes directly to console via _get_console(), never through logging.
    Shows multiple tasks if parallel operations are active, otherwise shows single task.
    """
    global _status_line_active

    if not _spinner_enabled:
        return

    # Determine what to display: parallel tasks take precedence over single task
    if _active_tasks:
        # Multiple parallel tasks - show them combined
        now = time.monotonic()
        # Find longest running task for elapsed time
        oldest_start = min(_active_tasks.values())
        elapsed = now - oldest_start
        elapsed_str = _format_elapsed(elapsed)

        # Build combined task list (truncate if too many)
        task_names = list(_active_tasks.keys())
        if len(task_names) == 1:
            display_tasks = task_names[0]
        elif len(task_names) <= 3:
            display_tasks = ", ".join(task_names)
        else:
            # Show first 2 + count of remaining
            display_tasks = f"{task_names[0]}, {task_names[1]} +{len(task_names) - 2} more"

        frame = _SPINNER_FRAMES[_spinner_idx % len(_SPINNER_FRAMES)]
        status = f"{frame} {display_tasks}... ({elapsed_str})"
    elif _current_task is not None:
        # Single task mode
        frame = _SPINNER_FRAMES[_spinner_idx % len(_SPINNER_FRAMES)]
        elapsed = time.monotonic() - _task_start_time
        elapsed_str = _format_elapsed(elapsed)
        status = f"{frame} {_current_task}... ({elapsed_str})"
    else:
        return

    # Truncate if too long for terminal
    width = _get_terminal_width()
    if len(status) > width - 1:
        status = status[: width - 4] + "..."

    # Write DIRECTLY to console, bypassing any logging redirections
    console = _get_console()
    try:
        console.write(f"\r\x1b[2K\x1b[1;36m{status}\x1b[0m")
        console.flush()
        _status_line_active = True
    except (OSError, ValueError):
        # Console might be closed/invalid, silently ignore
        pass


def _clear_status_line() -> None:
    """Clear the status line content.

    IMPORTANT: Output goes directly to console via _get_console(), never through logging.
    """
    global _status_line_active

    if _status_line_active:
        console = _get_console()
        try:
            console.write("\r\x1b[2K")
            console.flush()
        except (OSError, ValueError):
            pass
        _status_line_active = False


def _finalize_status_line() -> None:
    """Finalize the status line - print newline so it stays visible.

    IMPORTANT: Output goes directly to console via _get_console(), never through logging.
    """
    global _status_line_active

    if _status_line_active:
        console = _get_console()
        try:
            console.write("\n")
            console.flush()
        except (OSError, ValueError):
            pass
        _status_line_active = False


def _spinner_loop() -> None:
    """Background thread loop that updates the spinner."""
    global _spinner_idx

    while _spinner_stop_event is not None and not _spinner_stop_event.is_set():
        with _spinner_lock:
            # Draw if we have any task (single or parallel)
            # Don't draw if there's a pending incomplete line (would overwrite it)
            has_task = _current_task is not None or len(_active_tasks) > 0
            if (
                has_task
                and _spinner_enabled
                and not _output_paused
                and not _pending_incomplete_line
            ):
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

    # Print cancellation message DIRECTLY to console (not through logging)
    console = _get_console()
    try:
        console.write("\n\033[33m[INTERRUPTED]\033[0m Ctrl+C received, terminating...\n")
        console.flush()
    except (OSError, ValueError):
        pass

    # Kill all tracked subprocesses and their children immediately
    _kill_process_trees()

    # Signal spinner thread to stop
    if _spinner_stop_event is not None:
        _spinner_stop_event.set()

    # Exit immediately - use os._exit to bypass any blocking cleanup
    os._exit(130)  # 128 + SIGINT(2)


def _kill_process_trees() -> None:
    """Kill all tracked subprocesses and their entire process trees."""
    with _process_lock:
        for proc in list(_active_processes):
            try:
                if proc.poll() is None:  # Still running
                    _kill_process_tree(proc.pid)
            except Exception:
                pass  # Best effort cleanup
        _active_processes.clear()


def _kill_process_tree(pid: int) -> None:
    """
    Kill a process and all its children/descendants.

    On Windows, uses taskkill /T for tree kill.
    On Unix, walks the process tree manually.
    """
    import platform

    if platform.system() == "Windows":
        # Use taskkill with /T to kill entire process tree, /F for force
        try:
            subprocess.run(
                ["taskkill", "/F", "/T", "/PID", str(pid)],
                capture_output=True,
                timeout=5,
            )
        except Exception:
            # Fallback: try to kill just the process
            try:
                os.kill(pid, signal.SIGTERM)
            except Exception:
                pass
    else:
        # Unix: kill process group if possible, otherwise just the process
        try:
            os.killpg(os.getpgid(pid), signal.SIGKILL)
        except (ProcessLookupError, PermissionError, OSError):
            try:
                os.kill(pid, signal.SIGKILL)
            except Exception:
                pass


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
    Kill all tracked subprocesses and their process trees.

    @return Number of processes that were killed.
    """
    killed = 0
    with _process_lock:
        for proc in list(_active_processes):
            try:
                if proc.poll() is None:  # Still running
                    _kill_process_tree(proc.pid)
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

    This resets the elapsed timer. Use update_task() to change the display
    text without resetting the timer.

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


def update_task(task_name: str) -> None:
    """
    Update the current task text WITHOUT resetting the elapsed timer.

    Use this for progress updates where you want to change the display
    but keep the same elapsed time counter. If no task is active, this
    acts like set_task() and starts a new timer.

    @param task_name The new task name to display.
    """
    global _current_task, _task_start_time, _spinner_idx

    with _spinner_lock:
        # If no task was active, start fresh with a timer
        if _current_task is None:
            _task_start_time = time.monotonic()
            _spinner_idx = 0
        # Just update the text, don't reset timer
        _current_task = task_name
        # Draw immediately
        _draw_status_line()


def clear_task() -> None:
    """Clear the current task (equivalent to set_task(None))."""
    set_task(None)


# ---------------------------------------------------------------------------
# Parallel task tracking - for showing multiple concurrent tasks
# ---------------------------------------------------------------------------


def add_parallel_task(task_name: str) -> None:
    """
    Add a parallel task to track. Multiple tasks can run concurrently.
    Use this for parallel operations instead of set_task().

    @param task_name The task name to track.
    """
    with _spinner_lock:
        if task_name not in _active_tasks:
            _active_tasks[task_name] = time.monotonic()
            # Draw immediately
            _draw_status_line()


def remove_parallel_task(task_name: str) -> None:
    """
    Remove a parallel task when it completes.

    @param task_name The task name to remove.
    """
    with _spinner_lock:
        _active_tasks.pop(task_name, None)
        # If no more tasks, clear the line
        if not _active_tasks and _current_task is None:
            _clear_status_line()


def clear_parallel_tasks() -> None:
    """Clear all parallel tasks."""
    with _spinner_lock:
        _active_tasks.clear()
        if _current_task is None:
            _clear_status_line()


def get_parallel_task_count() -> int:
    """Get the number of currently active parallel tasks."""
    return len(_active_tasks)


def pause_for_output() -> None:
    """
    Prepare for log output by clearing spinner line.
    Call resume_after_output() when done printing.
    """
    global _output_paused, _status_line_active
    with _spinner_lock:
        if not _output_paused:
            if _status_line_active:
                # Clear the spinner line - log will print here
                try:
                    console = _get_console()
                    console.write("\r\x1b[2K")
                    console.flush()
                except Exception:
                    pass
                _status_line_active = False
            _output_paused = True


def resume_after_output() -> None:
    """
    Redraw the status line after log output is complete.
    The log output ends with newline, so spinner goes on the next line.
    Does NOT redraw if there's a pending incomplete line.
    """
    global _output_paused
    with _spinner_lock:
        if _output_paused:
            _output_paused = False
            # Don't draw if there's a pending incomplete line (would overwrite it)
            if _pending_incomplete_line:
                return
            # Draw if we have any task (single or parallel)
            has_task = _current_task is not None or len(_active_tasks) > 0
            if has_task and _spinner_enabled:
                _draw_status_line()


def mark_incomplete_line() -> None:
    """
    Mark that there's an incomplete line awaiting continuation.
    This prevents the spinner from overwriting the pending output.
    Call clear_incomplete_line() after completing the line.
    """
    global _pending_incomplete_line
    with _spinner_lock:
        _pending_incomplete_line = True


def clear_incomplete_line() -> None:
    """
    Clear the incomplete line flag after completing the line.
    This allows the spinner to resume drawing.
    """
    global _pending_incomplete_line
    with _spinner_lock:
        _pending_incomplete_line = False


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
