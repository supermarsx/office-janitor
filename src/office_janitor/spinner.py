"""!
@brief Global persistent spinner for showing current task status.
@details Provides a spinner that always shows the current operation at the
bottom of the console output. The spinner redraws after log lines and updates
to show the current task name with elapsed time.

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

# Spinner animation frames - use ASCII-safe characters for Windows compatibility
# Traditional spinner: | / - \
_SPINNER_FRAMES = ("|", "/", "-", "\\")

# Global spinner state
_spinner_lock = threading.Lock()
_current_task: str | None = None
_task_start_time: float = 0.0
_spinner_idx: int = 0
_last_line_len: int = 0
_spinner_enabled: bool = True
_spinner_thread: threading.Thread | None = None
_spinner_stop_event: threading.Event | None = None
_output_in_progress: bool = False  # Flag to suppress spinner during rapid output

# Track active subprocesses for cleanup on SIGINT
# Use Any for Popen type param since it's constrained to bytes|str
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


def _clear_line() -> None:
    """Clear the current spinner line from console."""
    global _last_line_len
    if _last_line_len > 0:
        sys.stdout.write(f"\r{' ' * _last_line_len}\r")
        sys.stdout.flush()
        _last_line_len = 0


def _draw_spinner() -> None:
    """Draw the spinner line showing current task."""
    global _spinner_idx, _last_line_len
    
    if not _spinner_enabled or _current_task is None:
        return
    
    frame = _SPINNER_FRAMES[_spinner_idx % len(_SPINNER_FRAMES)]
    _spinner_idx += 1
    
    elapsed = time.monotonic() - _task_start_time
    elapsed_str = _format_elapsed(elapsed)
    
    line = f"{frame} Working on: {_current_task}... ({elapsed_str})"
    
    sys.stdout.write(f"\r{line}")
    sys.stdout.flush()
    _last_line_len = len(line)


def _spinner_loop() -> None:
    """Background thread loop that updates the spinner."""
    global _output_in_progress
    while _spinner_stop_event is not None and not _spinner_stop_event.is_set():
        with _spinner_lock:
            # Only draw if there's a task and no output is happening
            if _current_task is not None and not _output_in_progress:
                _draw_spinner()
            # Reset the output flag - if no new output in this tick, we can draw next time
            _output_in_progress = False
        _spinner_stop_event.wait(0.1)  # Update every 100ms


# ---------------------------------------------------------------------------
# SIGINT / Ctrl+C handling
# ---------------------------------------------------------------------------


def _sigint_handler(signum: int, frame: object) -> None:
    """
    Handle SIGINT (Ctrl+C) for instant termination.
    
    Clears the spinner, kills all tracked subprocesses, and exits immediately.
    """
    # Clear spinner line first
    with _spinner_lock:
        _clear_line()
    
    # Print cancellation message
    print("\n\033[33m[INTERRUPTED]\033[0m Ctrl+C received, terminating...", flush=True)
    
    # Kill all tracked subprocesses
    with _process_lock:
        if _active_processes:
            print(f"Terminating {len(_active_processes)} active subprocess(es)...", flush=True)
            for proc in list(_active_processes):
                try:
                    if proc.poll() is None:  # Still running
                        # On Windows, terminate() sends SIGTERM equivalent
                        proc.terminate()
                        try:
                            proc.wait(timeout=2.0)
                        except subprocess.TimeoutExpired:
                            # Force kill if terminate didn't work
                            proc.kill()
                            proc.wait(timeout=1.0)
                except Exception:
                    pass  # Best effort cleanup
            _active_processes.clear()
    
    # Stop spinner thread
    stop_spinner_thread()
    
    print("Exiting.", flush=True)
    
    # Exit with interrupted status code
    sys.exit(130)  # 128 + SIGINT(2)


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
    _spinner_thread = threading.Thread(
        target=_spinner_loop,
        daemon=True,
        name="spinner"
    )
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


def set_task(task_name: str | None) -> None:
    """
    Set the current task being worked on.
    
    @param task_name The task name to display, or None to clear.
    """
    global _current_task, _task_start_time, _spinner_idx
    
    with _spinner_lock:
        _clear_line()
        _current_task = task_name
        if task_name is not None:
            _task_start_time = time.monotonic()
            _spinner_idx = 0
            # Don't draw immediately - let the spinner thread handle it
            # This keeps all drawing in one place and prevents flicker


def clear_task() -> None:
    """Clear the current task (equivalent to set_task(None))."""
    set_task(None)


def pause_for_output() -> None:
    """
    Temporarily clear spinner for other output.
    Call resume_after_output() after printing.
    """
    global _output_in_progress
    with _spinner_lock:
        _clear_line()
        _output_in_progress = True  # Signal that output is happening


def resume_after_output() -> None:
    """
    Signal that output is done. Spinner will redraw on its next tick.
    Does NOT immediately redraw - lets spinner thread handle it for cleaner output.
    """
    # Don't redraw here - let the spinner thread handle it on its next 100ms tick
    # This prevents spinner from appearing between rapid log lines
    pass


def spinner_print(message: str, **kwargs: object) -> None:
    """
    Print a message while preserving the spinner.
    Clears spinner, prints message. Spinner redraws on next tick.
    """
    global _output_in_progress
    with _spinner_lock:
        _clear_line()
        _output_in_progress = True
        print(message, **kwargs)


def enable_spinner(enabled: bool = True) -> None:
    """Enable or disable the spinner globally."""
    global _spinner_enabled
    with _spinner_lock:
        if not enabled:
            _clear_line()
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
        _clear_line()
    stop_spinner_thread()


atexit.register(_cleanup_spinner)
