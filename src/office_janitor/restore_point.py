"""!
@brief System restore point management.
@details Creates restore points prior to destructive operations when supported,
providing optional rollback coverage per the specification.
"""
from __future__ import annotations

import subprocess

from . import logging_ext


def create_restore_point(description: str) -> None:
    """!
    @brief Request a system restore point with the supplied description.
    """
    human_logger = logging_ext.get_human_logger()

    command = [
        "powershell.exe",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        (
            "Checkpoint-Computer -Description \"{}\" "
            "-RestorePointType 'MODIFY_SETTINGS'".format(description.replace("\"", "'"))
        ),
    ]

    human_logger.info("Creating system restore point: %s", description)

    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=120,
            check=False,
        )
    except FileNotFoundError:  # pragma: no cover - Windows-specific command.
        human_logger.debug("powershell.exe unavailable; skipping restore point creation")
        return

    if result.returncode != 0:
        human_logger.warning(
            "Restore point creation returned %s: %s",
            result.returncode,
            result.stderr.strip(),
        )
