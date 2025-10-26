"""!
@brief Click-to-Run uninstall orchestration utilities.
@details Mirrors the Click-to-Run OffScrub flow by composing the same VBS helper
invocations as the reference ``OfficeScrubber.cmd`` script while retaining the
project's structured logging and dry-run guarantees.
"""
from __future__ import annotations

import subprocess
import time
from pathlib import Path
from typing import Iterable, List, Mapping, Sequence

from . import constants, logging_ext

OFFSCRUB_C2R_ARGS = constants.C2R_OFFSCRUB_ARGS
"""!
@brief Backwards-compatible alias for Click-to-Run OffScrub arguments.
"""
from .off_scrub_scripts import ensure_offscrub_script

C2R_TIMEOUT = 3600
"""!
@brief Timeout (seconds) for Click-to-Run removal operations.
"""


def _collect_release_ids(raw: Iterable[str] | Sequence[str] | str | None) -> List[str]:
    """!
    @brief Normalise Click-to-Run release identifiers into a list.
    """

    if raw is None:
        return []
    if isinstance(raw, str):
        return [item.strip() for item in raw.split(",") if item.strip()]
    return [str(item).strip() for item in raw if str(item).strip()]


def build_command(
    config: Mapping[str, object],
    *,
    script_directory: Path | None = None,
) -> List[str]:
    """!
    @brief Compose the OffScrub Click-to-Run command for the given inventory entry.
    """

    release_ids = _collect_release_ids(
        config.get("release_ids")
        or config.get("products")
        or config.get("ProductReleaseIds")
    )

    script_path = ensure_offscrub_script(
        constants.C2R_OFFSCRUB_SCRIPT, base_directory=script_directory
    )

    command: List[str] = [str(constants.OFFSCRUB_EXECUTABLE)]
    command.extend(str(arg) for arg in constants.OFFSCRUB_HOST_ARGS)
    command.append(str(script_path))
    command.extend(constants.C2R_OFFSCRUB_ARGS)
    if release_ids:
        command.append(f"/PRODUCTS={';'.join(release_ids)}")
    return command


def uninstall_products(config: Mapping[str, object], *, dry_run: bool = False) -> None:
    """!
    @brief Trigger Click-to-Run OffScrub helpers for the supplied configuration.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    command = build_command(config)
    release_ids = _collect_release_ids(config.get("release_ids") or config.get("products"))

    machine_logger.info(
        "c2r_uninstall_plan",
        extra={
            "event": "c2r_uninstall_plan",
            "release_ids": release_ids or None,
            "dry_run": bool(dry_run),
            "command": command,
        },
    )

    if dry_run:
        human_logger.info("Dry-run: would invoke %s", " ".join(command))
        return

    human_logger.info(
        "Invoking OffScrubC2R helper for release identifiers: %s",
        ", ".join(release_ids) if release_ids else "ALL",
    )
    start = time.monotonic()
    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=C2R_TIMEOUT,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        duration = time.monotonic() - start
        human_logger.error(
            "OffScrubC2R timed out after %.1fs for releases %s",
            duration,
            ", ".join(release_ids) if release_ids else "ALL",
        )
        machine_logger.error(
            "c2r_uninstall_timeout",
            extra={
                "event": "c2r_uninstall_timeout",
                "release_ids": release_ids or None,
                "duration": duration,
                "stdout": exc.stdout,
                "stderr": exc.stderr,
            },
        )
        raise RuntimeError("Click-to-Run uninstall timed out") from exc

    duration = time.monotonic() - start
    if result.returncode != 0:
        human_logger.error(
            "OffScrubC2R failed with exit code %s for releases %s",
            result.returncode,
            ", ".join(release_ids) if release_ids else "ALL",
        )
        machine_logger.error(
            "c2r_uninstall_failure",
            extra={
                "event": "c2r_uninstall_failure",
                "release_ids": release_ids or None,
                "return_code": result.returncode,
                "stdout": result.stdout,
                "stderr": result.stderr,
                "duration": duration,
            },
        )
        raise RuntimeError("Click-to-Run uninstall failed")

    human_logger.info(
        "Successfully completed OffScrubC2R in %.1f seconds for releases %s.",
        duration,
        ", ".join(release_ids) if release_ids else "ALL",
    )
    machine_logger.info(
        "c2r_uninstall_success",
        extra={
            "event": "c2r_uninstall_success",
            "release_ids": release_ids or None,
            "return_code": result.returncode,
            "stdout": result.stdout,
            "stderr": result.stderr,
            "duration": duration,
        },
    )
