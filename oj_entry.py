"""!
@brief Shim entry point for Office Janitor.
@details This module ensures the package in ``src/`` is importable before
transferring control to :func:`office_janitor.main.main`.
"""

from __future__ import annotations

import os
import sys
import time

__all__ = ["main"]

# ---------------------------------------------------------------------------
# Startup timing and progress logging
# ---------------------------------------------------------------------------

_STARTUP_TIME = time.perf_counter()


def _elapsed_secs() -> float:
    """Return seconds since startup."""
    return time.perf_counter() - _STARTUP_TIME


def _log_init(message: str, *, newline: bool = True) -> None:
    """Print an initialization progress message with dmesg-style timestamp."""
    timestamp = f"[{_elapsed_secs():12.6f}]"
    if newline:
        print(f"{timestamp} {message}", flush=True)
    else:
        print(f"{timestamp} {message}", end="", flush=True)


def _log_ok(extra: str = "") -> None:
    """Print OK status in Linux init style [  OK  ]."""
    suffix = f" {extra}" if extra else ""
    print(f" [  \033[32mOK\033[0m  ]{suffix}", flush=True)


def _log_fail(reason: str = "") -> None:
    """Print FAIL status in Linux init style [FAILED]."""
    suffix = f" ({reason})" if reason else ""
    print(f"\r[\033[31mFAILED\033[0m]{suffix}", flush=True)


# ---------------------------------------------------------------------------
# Environment detection
# ---------------------------------------------------------------------------


def _is_frozen() -> bool:
    """Check if running as a PyInstaller bundle."""
    return hasattr(sys, "_MEIPASS")


def _get_bundle_path() -> str:
    """Get the PyInstaller bundle extraction path, or empty string if not frozen."""
    return getattr(sys, "_MEIPASS", "")


# ---------------------------------------------------------------------------
# Path setup
# ---------------------------------------------------------------------------

_REPO_ROOT = os.path.dirname(os.path.abspath(__file__))
_SRC_PATH = os.path.join(_REPO_ROOT, "src")
_PACKAGE_PATH = os.path.join(_SRC_PATH, "office_janitor")

_log_init("Office Janitor initializing...")
_log_init(f"  Executable: {sys.executable}")
_log_init(f"  Python version: {sys.version.split()[0]}")
_log_init(f"  Platform: {sys.platform}")
_log_init(f"  Frozen: {_is_frozen()}")

if _is_frozen():
    _log_init(f"  Bundle path: {_get_bundle_path()}")
else:
    _log_init(f"  Repository root: {_REPO_ROOT}")
    _log_init(f"  Source path: {_SRC_PATH}")

_log_init("Configuring module search paths...", newline=False)
if os.path.isdir(_SRC_PATH) and _SRC_PATH not in sys.path:
    sys.path.insert(0, _SRC_PATH)
_log_ok()

# Ensure the package is collected by PyInstaller
_log_init("Importing office_janitor package...", newline=False)
try:
    import office_janitor

    _log_ok()
except ImportError as e:
    _log_fail(str(e))
    raise

_log_init("Importing office_janitor.main module...", newline=False)
try:
    import office_janitor.main

    _log_ok()
except ImportError as e:
    _log_fail(str(e))
    raise

if os.path.isdir(_PACKAGE_PATH):
    __path__ = [_PACKAGE_PATH]
    if __spec__ is not None:  # pragma: no cover - import system attribute
        __spec__.submodule_search_locations = list(__path__)


def _prepend_src_to_sys_path() -> None:
    """!
    @brief Prepend the repository ``src`` directory to ``sys.path``.
    @details The shim mirrors the structure described in :mod:`spec.md`, keeping
    the distributable executable simple while letting the package live under
    ``src/``. In PyInstaller bundles, ``src`` is already in ``sys.path``.
    """

    # In PyInstaller bundles, _MEIPASS is set and src is already in sys.path
    if _is_frozen():
        return

    if _SRC_PATH not in sys.path:
        sys.path.insert(0, _SRC_PATH)


def _load_submodules() -> dict[str, bool]:
    """!
    @brief Pre-load critical submodules and report status.
    @returns Dictionary mapping module names to load success status.
    """
    submodules = [
        "office_janitor.version",
        "office_janitor.constants",
        "office_janitor.logging_ext",
        "office_janitor.elevation",
        "office_janitor.detect",
        "office_janitor.plan",
        "office_janitor.scrub",
        "office_janitor.ui",
    ]
    results = {}
    for mod_name in submodules:
        short_name = mod_name.split(".")[-1]
        _log_init(f"  Loading {short_name}...", newline=False)
        try:
            __import__(mod_name)
            _log_ok()
            results[mod_name] = True
        except ImportError as e:
            _log_fail(str(e))
            results[mod_name] = False
    return results


def main() -> int:
    """!
    @brief Invoke the package entry point after preparing ``sys.path``.
    @returns Exit status propagated from :func:`office_janitor.main.main`.
    """
    _log_init("Preparing environment...")
    _prepend_src_to_sys_path()

    _log_init("Pre-loading submodules...")
    load_results = _load_submodules()
    loaded = sum(1 for v in load_results.values() if v)
    total = len(load_results)
    _log_init(f"  Loaded {loaded}/{total} submodules")

    _log_init("Importing main entry point...", newline=False)
    try:
        from office_janitor.main import main as package_main

        _log_ok()
    except ImportError as e:
        _log_fail(str(e))
        raise

    # Try to get version info for logging
    try:
        from office_janitor.version import build_info

        info = build_info()
        _log_init(
            f"Version: {info.get('version', 'unknown')} (build {info.get('build', 'unknown')})"
        )
    except Exception:
        pass

    _log_init(f"Initialization complete in {_elapsed_secs():.3f}s")
    _log_init("-" * 60)

    # Pass startup time to main for continuous timestamps
    return package_main(start_time=_STARTUP_TIME)


if __name__ == "__main__":  # pragma: no cover - manual invocation
    sys.exit(main())

# Delegate key attributes to the real package module so importing this shim
# behaves like importing ``office_janitor.main`` (for tests/tooling that import
# the shim path first).
try:
    import importlib

    _prepend_src_to_sys_path()
    _package_main = importlib.import_module("office_janitor.main")
    ensure_admin_and_relaunch_if_needed = getattr(
        _package_main, "ensure_admin_and_relaunch_if_needed", None
    )
    enable_vt_mode_if_possible = getattr(_package_main, "enable_vt_mode_if_possible", None)
    build_arg_parser = getattr(_package_main, "build_arg_parser", None)
    _determine_mode = getattr(_package_main, "_determine_mode", None)
    _collect_plan_options = getattr(_package_main, "_collect_plan_options", None)
except Exception:  # pragma: no cover - best-effort delegation
    _package_main = None
