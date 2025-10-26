"""!
@brief Filesystem utilities for Office residue cleanup.
@details Future implementations discover install footprints, reset ACLs, and
remove leftovers from program directories and user profiles, matching the
specification requirements.
"""
from __future__ import annotations

import os
import shutil
import stat
import subprocess
from pathlib import Path
from typing import Iterable, Sequence

from . import logging_ext


def _handle_readonly(function, path: str, exc_info) -> None:  # pragma: no cover - defensive callback
    """!
    @brief Clear read-only attributes before retrying removal.
    """

    if isinstance(exc_info[1], PermissionError):
        os.chmod(path, stat.S_IWRITE)
        function(path)
    else:
        raise exc_info[1]


def remove_paths(paths: Iterable[Path], *, dry_run: bool = False) -> None:
    """!
    @brief Delete the supplied paths recursively while respecting dry-run behavior.
    """
    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    for raw in paths:
        target = Path(raw)
        machine_logger.info(
            "filesystem_remove_plan",
            extra={
                "event": "filesystem_remove_plan",
                "path": str(target),
                "dry_run": bool(dry_run),
            },
        )

        if dry_run:
            human_logger.info("Dry-run: would remove %s", target)
            continue

        try:
            make_paths_writable([target])
            reset_acl(target)
        except Exception as exc:  # pragma: no cover - logged for diagnostics
            human_logger.warning("Unable to reset ACLs for %s: %s", target, exc)

        if not target.exists():
            human_logger.debug("Skipping %s because it does not exist", target)
            continue

        human_logger.info("Removing %s", target)
        if target.is_dir():
            shutil.rmtree(target, onerror=_handle_readonly)
        else:
            try:
                target.unlink()
            except PermissionError:
                os.chmod(target, stat.S_IWRITE)
                target.unlink()


def reset_acl(path: Path) -> None:
    """!
    @brief Reset permissions on ``path`` so cleanup operations can proceed.
    """
    human_logger = logging_ext.get_human_logger()

    command = [
        "icacls",
        str(path),
        "/reset",
        "/t",
        "/c",
    ]

    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except FileNotFoundError:  # pragma: no cover - occurs on non-Windows hosts.
        human_logger.debug("icacls is not available; skipping ACL reset for %s", path)
        return

    if result.returncode != 0:
        human_logger.warning(
            "icacls reported exit code %s for %s: %s",
            result.returncode,
            path,
            result.stderr.strip(),
        )


def make_paths_writable(paths: Sequence[Path], *, dry_run: bool = False) -> None:
    """!
    @brief Clear read-only attributes in preparation for recursive deletion.
    @details Mirrors the ``attrib -R`` behaviour from ``OfficeScrubberAIO.cmd`` to
    ensure licensing and Click-to-Run directories can be removed regardless of
    inherited ACLs.
    """

    human_logger = logging_ext.get_human_logger()

    for raw in paths:
        target = Path(raw)
        if dry_run:
            human_logger.info("Dry-run: would clear attributes for %s", target)
            continue

        for suffix in ("", "\\*"):
            command = [
                "attrib.exe",
                "-R",
                f"{target}{suffix}",
                "/S",
                "/D",
            ]
            try:
                result = subprocess.run(
                    command,
                    capture_output=True,
                    text=True,
                    timeout=60,
                    check=False,
                )
            except FileNotFoundError:  # pragma: no cover - non-Windows host.
                human_logger.debug("attrib.exe unavailable; skipping attribute reset for %s", target)
                break

            if result.returncode not in {0, 1}:  # attrib returns 1 when no files match
                human_logger.debug(
                    "attrib exited with %s for %s: %s",
                    result.returncode,
                    target,
                    result.stderr.strip(),
                )
                break
