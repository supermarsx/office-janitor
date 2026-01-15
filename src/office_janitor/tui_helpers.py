"""!
@file tui_helpers.py
@brief Helper functions for the Office Janitor TUI.

@details Provides standalone utility functions for ANSI support detection,
key decoding, screen management, progress bar rendering, inventory formatting,
plan formatting, and input handling. These helpers are used by the main TUI
controller but have no dependencies on TUI state.
"""

from __future__ import annotations

import os
import re
import sys
import time
from collections.abc import Mapping
from typing import Any

from . import plan as plan_module

try:  # pragma: no cover - Windows specific
    import msvcrt as _msvcrt
except ImportError:  # pragma: no cover - non-Windows hosts
    _msvcrt = None

msvcrt: Any = _msvcrt


# ---------------------------------------------------------------------------
# ANSI support and terminal utilities
# ---------------------------------------------------------------------------


def supports_ansi(stream: object | None = None) -> bool:
    """!
    @brief Determine whether the current stdout supports ANSI escape sequences.
    """

    target = stream if stream is not None else sys.stdout
    if not hasattr(target, "isatty"):
        return False
    try:
        if not target.isatty():
            return False
    except Exception:
        return False

    if os.name != "nt":
        return True
    return bool(
        os.environ.get("WT_SESSION")
        or os.environ.get("ANSICON")
        or os.environ.get("TERM_PROGRAM")
        or os.environ.get("ConEmuANSI")
    )


def clear_screen() -> None:
    """Clear the terminal screen and move cursor to top-left."""
    sys.stdout.write("\x1b[2J\x1b[H")
    sys.stdout.flush()


def strip_ansi(text: str) -> str:
    """Remove ANSI escape sequences from text to get visible length."""
    return re.sub(r"\x1b\[[0-9;]*m", "", text)


def divider(width: int) -> str:
    """Create a horizontal divider line."""
    return "-" * width


# ---------------------------------------------------------------------------
# Key input handling
# ---------------------------------------------------------------------------


def decode_key(raw: str) -> str:
    """!
    @brief Convert raw key input into a normalized command token.
    """

    if not raw:
        return ""
    mapping = {
        "\t": "tab",
        "\r": "enter",
        "\n": "enter",
        " ": "space",
        "q": "quit",
        "Q": "quit",
        "\x1b": "escape",
        # vim-style hjkl navigation
        "h": "left",
        "j": "down",
        "k": "up",
        "l": "right",
        "H": "left",
        "J": "down",
        "K": "up",
        "L": "right",
    }
    if raw in mapping:
        return mapping[raw]
    if raw.startswith("\x1b"):
        sequences = {
            "\x1b[A": "up",
            "\x1b[B": "down",
            "\x1b[C": "right",
            "\x1b[D": "left",
            "\x1b[5~": "page_up",
            "\x1b[6~": "page_down",
            "\x1bOP": "f1",
            "\x1b[21~": "f10",
            "\x1b[1;9A": "up",
            "\x1b[1;9B": "down",
        }
        return sequences.get(raw, "")
    windows_sequences = {
        "\x00H": "up",
        "\x00P": "down",
        "\x00K": "left",
        "\x00M": "right",
        "\xe0H": "up",
        "\xe0P": "down",
        "\xe0K": "left",
        "\xe0M": "right",
        "\x00I": "page_up",
        "\x00Q": "page_down",
        "\xe0I": "page_up",
        "\xe0Q": "page_down",
        "\x00;": "f1",
        "\x00h": "f10",
    }
    if raw in windows_sequences:
        return windows_sequences[raw]
    trimmed = raw.strip()
    if trimmed == "/":
        return "/"
    return trimmed


def default_key_reader() -> str:
    """!
    @brief Read a key press using platform-specific methods.
    @returns Raw key string or empty string if no key pressed.
    """
    try:  # pragma: no cover - depends on Windows availability
        import msvcrt

        if msvcrt.kbhit():
            first = msvcrt.getwch()
            if first in {"\x00", "\xe0"}:
                second = msvcrt.getwch()
                return first + second
            if first == "\x1b" and msvcrt.kbhit():
                second = msvcrt.getwch()
                if second == "[":
                    third = msvcrt.getwch()
                    return first + second + third
            return first
        return ""
    except Exception:
        return read_input_line("Command (arrows navigate, enter selects, q to quit): ").strip()


def read_input_line(prompt: str) -> str:
    """Read a line of input from the user."""
    return input(prompt)


# ---------------------------------------------------------------------------
# Progress indicators
# ---------------------------------------------------------------------------


def spinner(duration: float, message: str) -> None:
    """!
    @brief Display a spinner animation for the specified duration.
    """
    frames = ["-", "\\", "|", "/"]
    start = time.monotonic()
    index = 0
    while time.monotonic() - start < duration:
        sys.stdout.write(f"\r{message} {frames[index % len(frames)]}")
        sys.stdout.flush()
        time.sleep(0.1)
        index += 1
    sys.stdout.write("\r" + " " * (len(message) + 2) + "\r")


def render_progress_bar(
    current: int,
    total: int,
    width: int = 40,
    fill_char: str = "█",
    empty_char: str = "░",
    show_percentage: bool = True,
) -> str:
    """!
    @brief Render a text-based progress bar with optional percentage.
    @param current Current progress value.
    @param total Total value (when current == total, bar is full).
    @param width Width of the bar in characters (excluding percentage).
    @param fill_char Character used for completed portion.
    @param empty_char Character used for incomplete portion.
    @param show_percentage If True, append percentage to bar.
    @returns Formatted progress bar string.
    """
    if total <= 0:
        percentage = 0.0
    else:
        percentage = min(100.0, (current / total) * 100.0)

    filled_width = int((percentage / 100.0) * width)
    empty_width = width - filled_width

    bar = f"{fill_char * filled_width}{empty_char * empty_width}"

    if show_percentage:
        return f"[{bar}] {percentage:5.1f}%"
    return f"[{bar}]"


# ---------------------------------------------------------------------------
# Inventory formatting
# ---------------------------------------------------------------------------


def format_inventory(inventory: Mapping[str, object]) -> list[str]:
    """!
    @brief Format an inventory mapping for display.
    """
    lines: list[str] = []
    for key, value in inventory.items():
        lines.extend(_flatten_inventory_entry(str(key), value))
    if not lines:
        lines.append("No data")
    return lines


def _flatten_inventory_entry(prefix: str, value: object) -> list[str]:
    """Recursively flatten an inventory entry for display."""
    if isinstance(value, Mapping):
        if not value:
            return [f"{prefix}: (empty)"]
        lines: list[str] = []
        for subkey, subvalue in value.items():
            nested_prefix = f"{prefix}.{subkey}"
            lines.extend(_flatten_inventory_entry(nested_prefix, subvalue))
        return lines
    if isinstance(value, (list, tuple, set)):
        items = list(value)
        if not items:
            return [f"{prefix}: (empty)"]
        return [f"{prefix}: {_stringify_inventory_value(item)}" for item in items]
    return [f"{prefix}: {_stringify_inventory_value(value)}"]


def _stringify_inventory_value(value: object) -> str:
    """Convert an inventory value to a string representation."""
    if isinstance(value, Mapping):
        if not value:
            return "{}"
        parts = [f"{key}={_stringify_inventory_value(subvalue)}" for key, subvalue in value.items()]
        return ", ".join(parts)
    if isinstance(value, (list, tuple, set)):
        if not value:
            return "[]"
        return ", ".join(_stringify_inventory_value(item) for item in value)
    return str(value)


def summarize_inventory(inventory: Mapping[str, object]) -> dict[str, int]:
    """!
    @brief Create a summary of inventory counts by category.
    """
    summary: dict[str, int] = {}
    for key, items in inventory.items():
        if hasattr(items, "__len__"):
            count = len(items)
        else:
            try:
                count = len(list(items))
            except TypeError:
                count = 1
        summary[str(key)] = count
    return summary


# ---------------------------------------------------------------------------
# Plan formatting
# ---------------------------------------------------------------------------


def format_plan(plan_data: list[dict[str, object]] | None) -> list[str]:
    """!
    @brief Format plan data for display in the TUI.
    """
    if not plan_data:
        return ["Plan not created"]

    summary = plan_module.summarize_plan(plan_data)
    lines: list[str] = [
        f"Steps: {summary.get('total_steps', 0)} (actionable {summary.get('actionable_steps', 0)})",
    ]

    mode = summary.get("mode")
    dry_run = summary.get("dry_run")
    if mode:
        lines.append(f"Mode: {mode}{' [dry-run]' if dry_run else ''}")

    target_versions = summary.get("target_versions") or []
    if target_versions:
        lines.append("Targets: " + ", ".join(str(item) for item in target_versions))

    discovered = summary.get("discovered_versions") or []
    if discovered:
        lines.append("Detected: " + ", ".join(str(item) for item in discovered))

    uninstall_versions = summary.get("uninstall_versions") or []
    if uninstall_versions:
        lines.append("Uninstalls: " + ", ".join(str(item) for item in uninstall_versions))

    cleanup_categories = summary.get("cleanup_categories") or []
    if cleanup_categories:
        lines.append("Cleanup: " + ", ".join(str(item) for item in cleanup_categories))

    requested_components = summary.get("requested_components") or []
    if requested_components:
        lines.append("Include: " + ", ".join(str(item) for item in requested_components))

    unsupported_components = summary.get("unsupported_components") or []
    if unsupported_components:
        lines.append(
            "Unsupported include: " + ", ".join(str(item) for item in unsupported_components)
        )

    categories = summary.get("categories") or {}
    if categories:
        formatted = ", ".join(f"{key}={value}" for key, value in categories.items())
        lines.append("Categories: " + formatted)

    return lines
