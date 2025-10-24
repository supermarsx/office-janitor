"""!
@brief Helpers for orchestrating MSI-based Office uninstalls.
@details This module locates MSI product codes, drives ``msiexec`` with the
correct flags, monitors progress, and captures logs according to the
specification.
"""
from __future__ import annotations

import subprocess
import time
from pathlib import Path
from typing import Iterable, List

from . import logging_ext

MSIEXEC = "msiexec.exe"
"""!
@brief Executable used to drive MSI uninstallations.
"""

MSIEXEC_TIMEOUT = 1800
"""!
@brief Default timeout (seconds) for ``msiexec`` executions.
"""


def _sanitise_product_code(product_code: str) -> str:
    """!
    @brief Produce a filesystem-safe representation of ``product_code``.
    @details ``msiexec`` product codes are GUIDs wrapped in braces. The helper
    strips these braces and collapses whitespace so the value can be used in log
    filenames without additional quoting.
    """

    return product_code.strip().strip("{}").replace("-", "").upper() or "unknown"


def _log_target_path(directory: Path | None, product_code: str) -> Path | None:
    """!
    @brief Resolve the log path for a given product code if logging is enabled.
    """

    if directory is None:
        return None
    safe_code = _sanitise_product_code(product_code)
    return directory / f"msi-{safe_code}.log"


def uninstall_products(product_codes: Iterable[str], *, dry_run: bool = False) -> None:
    """!
    @brief Uninstall the supplied MSI product codes while respecting dry-run semantics.
    """
    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()
    log_directory = logging_ext.get_log_directory()

    codes: List[str] = [code for code in (str(code) for code in product_codes) if code.strip()]
    if not codes:
        human_logger.info("No MSI product codes supplied for uninstall; skipping.")
        return

    human_logger.info(
        "Preparing to uninstall %d MSI product(s). Dry-run: %s", len(codes), bool(dry_run)
    )
    failures: List[str] = []

    for product_code in codes:
        command: List[str] = [MSIEXEC, "/x", product_code, "/qb!", "/norestart"]
        log_path = _log_target_path(log_directory, product_code)
        if log_path is not None:
            command.extend(["/log", str(log_path)])

        machine_logger.info(
            "msi_uninstall_plan",
            extra={
                "event": "msi_uninstall_plan",
                "product_code": product_code,
                "dry_run": bool(dry_run),
                "log_path": str(log_path) if log_path else None,
            },
        )

        if dry_run:
            human_logger.info("Dry-run: would invoke %s", " ".join(command))
            continue

        human_logger.info("Invoking msiexec for %s", product_code)
        start = time.monotonic()
        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=MSIEXEC_TIMEOUT,
                check=False,
            )
        except subprocess.TimeoutExpired as exc:
            duration = time.monotonic() - start
            message = (
                f"msiexec timed out after {duration:.1f}s for {product_code}:"
                f" stdout={exc.stdout!r} stderr={exc.stderr!r}"
            )
            human_logger.error(message)
            machine_logger.error(
                "msi_uninstall_timeout",
                extra={
                    "event": "msi_uninstall_timeout",
                    "product_code": product_code,
                    "duration": duration,
                },
            )
            failures.append(product_code)
            continue

        duration = time.monotonic() - start
        if result.returncode != 0:
            human_logger.error(
                "msiexec for %s failed with exit code %s", product_code, result.returncode
            )
            machine_logger.error(
                "msi_uninstall_failure",
                extra={
                    "event": "msi_uninstall_failure",
                    "product_code": product_code,
                    "return_code": result.returncode,
                    "stdout": result.stdout,
                    "stderr": result.stderr,
                    "duration": duration,
                },
            )
            failures.append(product_code)
            continue

        human_logger.info(
            "Successfully uninstalled %s via msiexec in %.1f seconds.", product_code, duration
        )
        machine_logger.info(
            "msi_uninstall_success",
            extra={
                "event": "msi_uninstall_success",
                "product_code": product_code,
                "return_code": result.returncode,
                "stdout": result.stdout,
                "stderr": result.stderr,
                "duration": duration,
            },
        )

    if failures:
        raise RuntimeError(f"Failed to uninstall MSI products: {', '.join(failures)}")
