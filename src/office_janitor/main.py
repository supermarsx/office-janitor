"""!
@brief Primary entry point for the Office Janitor CLI.
@details This module bootstraps the runtime environment described in
:mod:`spec.md`: argument parsing for detection and scrubbing modes, validating
administrative elevation, enabling Windows VT mode when available, and invoking
logging setup so future sub-systems can emit structured telemetry.
"""
from __future__ import annotations

import argparse
import ctypes
import logging
import os
import pathlib
import sys
from typing import Iterable, Optional

from . import logging_ext, version


def enable_vt_mode_if_possible() -> None:
    """!
    @brief Attempt to enable ANSI/VT processing on Windows consoles.
    @details Per the specification, the application should try to enable virtual
    terminal support so both the plain CLI and future TUI renderer can emit
    colorized output. Failures are silently ignored because the feature is
    optional.
    """

    try:
        from ctypes import wintypes

        kernel32 = ctypes.windll.kernel32  # type: ignore[attr-defined]
        handle = kernel32.GetStdHandle(-11)  # STD_OUTPUT_HANDLE
        mode = wintypes.DWORD()
        if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
            kernel32.SetConsoleMode(handle, mode.value | 0x0004)
    except Exception:
        # Non-Windows platforms or consoles without VT capability are fine.
        return


def ensure_admin_and_relaunch_if_needed() -> None:
    """!
    @brief Request elevation if the current process lacks administrative rights.
    """

    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()  # type: ignore[attr-defined]
    except Exception:
        is_admin = False
    if not is_admin and os.name == "nt":
        params = " ".join(f'"{arg}"' for arg in sys.argv)
        ctypes.windll.shell32.ShellExecuteW(  # type: ignore[attr-defined]
            None, "runas", sys.executable, params, None, 1
        )
        sys.exit(0)


def build_arg_parser() -> argparse.ArgumentParser:
    """!
    @brief Create the top-level argument parser with the specification's surface area.
    @details The parser wires in the mutually exclusive modes and shared options
    defined in the specification so later feature work can hook into the parsed
    values without changing the public CLI signature.
    """

    parser = argparse.ArgumentParser(prog="office-janitor", add_help=True)
    parser.add_argument(
        "-V",
        "--version",
        action="version",
        version=f"{version.__version__} ({version.__build__})",
    )

    modes = parser.add_mutually_exclusive_group()
    modes.add_argument("--auto-all", action="store_true", help="Run full detection and scrub.")
    modes.add_argument(
        "--target",
        metavar="VER",
        help="Target a specific Office version (2003-2024/365).",
    )
    modes.add_argument("--diagnose", action="store_true", help="Emit inventory and plan without changes.")
    modes.add_argument("--cleanup-only", action="store_true", help="Skip uninstalls; clean residue and licensing.")

    parser.add_argument("--include", metavar="COMPONENTS", help="Additional suites/apps to include.")
    parser.add_argument("--force", action="store_true", help="Relax certain guardrails when safe.")
    parser.add_argument("--dry-run", action="store_true", help="Simulate actions without modifying the system.")
    parser.add_argument("--no-restore-point", action="store_true", help="Skip creating a restore point.")
    parser.add_argument("--no-license", action="store_true", help="Skip license cleanup steps.")
    parser.add_argument("--keep-templates", action="store_true", help="Preserve user templates like normal.dotm.")
    parser.add_argument("--plan", metavar="OUT", help="Write the computed action plan to a JSON file.")
    parser.add_argument("--logdir", metavar="DIR", help="Directory for human/JSONL log output.")
    parser.add_argument("--backup", metavar="DIR", help="Destination for registry/file backups.")
    parser.add_argument("--timeout", metavar="SEC", type=int, help="Per-step timeout in seconds.")
    parser.add_argument("--quiet", action="store_true", help="Minimal console output (errors only).")
    parser.add_argument("--json", action="store_true", help="Mirror structured events to stdout.")
    parser.add_argument("--tui", action="store_true", help="Force the interactive text UI mode.")
    parser.add_argument("--no-color", action="store_true", help="Disable ANSI color codes.")
    parser.add_argument("--tui-compact", action="store_true", help="Use a compact TUI layout for small consoles.")
    parser.add_argument(
        "--tui-refresh",
        metavar="MS",
        type=int,
        help="Refresh interval for the TUI renderer in milliseconds.",
    )
    return parser


def _resolve_log_directory(candidate: Optional[str]) -> pathlib.Path:
    """!
    @brief Determine the log directory path using specification defaults when unspecified.
    """

    if candidate:
        return pathlib.Path(candidate).expanduser().resolve()
    if os.name == "nt":
        program_data = os.environ.get("ProgramData", r"C:\\ProgramData")
        return pathlib.Path(program_data) / "OfficeJanitor" / "logs"
    return pathlib.Path.cwd() / "logs"


def _bootstrap_logging(args: argparse.Namespace) -> tuple[logging.Logger, logging.Logger]:
    """!
    @brief Initialize human and machine loggers, falling back if unimplemented.
    """

    logdir = _resolve_log_directory(getattr(args, "logdir", None))
    try:
        return logging_ext.setup_logging(logdir, json_to_stdout=getattr(args, "json", False))
    except NotImplementedError:
        logging.basicConfig(level=logging.INFO)
        human = logging.getLogger("human")
        machine = logging.getLogger("machine")
        return human, machine


def main(argv: Optional[Iterable[str]] = None) -> int:
    """!
    @brief Entry point invoked by the shim and PyInstaller bundle.
    """

    ensure_admin_and_relaunch_if_needed()
    enable_vt_mode_if_possible()
    parser = build_arg_parser()
    args = parser.parse_args(list(argv) if argv is not None else None)
    human_log, machine_log = _bootstrap_logging(args)

    human_log.info("Office Janitor bootstrap complete; core logic not yet implemented.")
    machine_log.info("startup", extra={"event": "startup", "data": {"mode": _determine_mode(args)}})
    return 0


def _determine_mode(args: argparse.Namespace) -> str:
    """!
    @brief Map parsed arguments to a simple textual mode identifier.
    """

    if getattr(args, "auto_all", False):
        return "auto-all"
    if getattr(args, "target", None):
        return f"target:{args.target}"
    if getattr(args, "diagnose", False):
        return "diagnose"
    if getattr(args, "cleanup_only", False):
        return "cleanup-only"
    return "interactive"


if __name__ == "__main__":  # pragma: no cover - for manual execution
    sys.exit(main())
