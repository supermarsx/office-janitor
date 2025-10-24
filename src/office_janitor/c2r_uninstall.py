"""!
@brief Click-to-Run uninstall orchestration utilities.
@details The routines invoke ``OfficeC2RClient.exe`` and related tools to remove
Click-to-Run Office releases while tracking progress and handling edge cases as
outlined in the specification.
"""
from __future__ import annotations

import subprocess
import time
from pathlib import Path
from typing import Iterable, List, Mapping

from . import logging_ext

DEFAULT_CLIENT_PATHS: tuple[Path, ...] = (
    Path(r"C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"),
    Path(r"C:\Program Files (x86)\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe"),
)
"""!
@brief Common search paths for ``OfficeC2RClient.exe``.
"""

CLIENT_TIMEOUT = 3600
"""!
@brief Timeout (seconds) for Click-to-Run uninstall operations.
"""


def _select_client_path(explicit: str | Path | None) -> Path:
    """!
    @brief Choose an ``OfficeC2RClient`` path using configuration hints.
    """

    if explicit:
        candidate = Path(explicit)
        return candidate

    for candidate in DEFAULT_CLIENT_PATHS:
        if candidate.exists():
            return candidate

    return DEFAULT_CLIENT_PATHS[0]


def _collect_release_ids(raw: Iterable[str] | str | None) -> List[str]:
    """!
    @brief Normalise release identifiers into a list for command construction.
    """

    if raw is None:
        return []
    if isinstance(raw, str):
        return [item.strip() for item in raw.split(",") if item.strip()]
    return [str(item).strip() for item in raw if str(item).strip()]


def uninstall_products(config: Mapping[str, str], *, dry_run: bool = False) -> None:
    """!
    @brief Trigger Click-to-Run uninstall sequences for the supplied configuration.
    """
    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    release_ids = _collect_release_ids(config.get("release_ids") or config.get("products"))
    if not release_ids:
        raise ValueError("Click-to-Run uninstall requires at least one release identifier")

    client_path = _select_client_path(config.get("client_path"))
    additional_args: Iterable[str] | None = config.get("additional_args")  # type: ignore[assignment]

    command: List[str] = [
        str(client_path),
        "/update",
        "user",
        f"displaylevel=false",
        "forceappshutdown=true",
        f"productstoremove={';'.join(release_ids)}",
        "productstoadd=none",
    ]

    if additional_args:
        command.extend(str(arg) for arg in additional_args)

    log_directory = logging_ext.get_log_directory()
    log_path: Path | None = None
    if log_directory is not None:
        joined = "-".join(release_ids)
        safe = joined.replace("/", "_").replace("\\", "_") or "c2r"
        log_path = log_directory / f"c2r-{safe}.log"
        command.extend(["/log", str(log_path)])

    machine_logger.info(
        "c2r_uninstall_plan",
        extra={
            "event": "c2r_uninstall_plan",
            "release_ids": release_ids,
            "client_path": str(client_path),
            "dry_run": bool(dry_run),
            "log_path": str(log_path) if log_path else None,
        },
    )

    if dry_run:
        human_logger.info("Dry-run: would invoke %s", " ".join(command))
        return

    start = time.monotonic()
    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=CLIENT_TIMEOUT,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        duration = time.monotonic() - start
        human_logger.error(
            "OfficeC2RClient timed out after %.1fs for releases %s", duration, ", ".join(release_ids)
        )
        machine_logger.error(
            "c2r_uninstall_timeout",
            extra={
                "event": "c2r_uninstall_timeout",
                "release_ids": release_ids,
                "duration": duration,
                "stdout": exc.stdout,
                "stderr": exc.stderr,
            },
        )
        raise RuntimeError("Click-to-Run uninstall timed out") from exc

    duration = time.monotonic() - start
    if result.returncode != 0:
        human_logger.error(
            "OfficeC2RClient failed with exit code %s for releases %s",
            result.returncode,
            ", ".join(release_ids),
        )
        machine_logger.error(
            "c2r_uninstall_failure",
            extra={
                "event": "c2r_uninstall_failure",
                "release_ids": release_ids,
                "return_code": result.returncode,
                "stdout": result.stdout,
                "stderr": result.stderr,
                "duration": duration,
            },
        )
        raise RuntimeError("Click-to-Run uninstall failed")

    human_logger.info(
        "OfficeC2RClient removed releases %s in %.1f seconds.", ", ".join(release_ids), duration
    )
    machine_logger.info(
        "c2r_uninstall_success",
        extra={
            "event": "c2r_uninstall_success",
            "release_ids": release_ids,
            "return_code": result.returncode,
            "stdout": result.stdout,
            "stderr": result.stderr,
            "duration": duration,
        },
    )
