"""!
@brief System restore point management utilities.
@details Coordinates PowerShell/WMI calls to create restore points prior to
running destructive operations. The helpers provide structured logging, dry-run
simulation, and defensive error handling so callers can safely request restore
coverage when available.
"""
from __future__ import annotations

import json
import os
import subprocess
import textwrap
from typing import Sequence

from . import logging_ext


_POWERSHELL_EXECUTABLE = "powershell.exe"
_POWERSHELL_TIMEOUT_SECONDS = 180


def create_restore_point(
    description: str,
    *,
    dry_run: bool = False,
    timeout: int = _POWERSHELL_TIMEOUT_SECONDS,
) -> bool:
    """!
    @brief Request a system restore point with the supplied description.
    @details Uses PowerShell to invoke the ``SystemRestore`` WMI provider, falling
    back to ``Checkpoint-Computer`` if the class is unavailable. When ``dry_run``
    is enabled the function emits simulation telemetry without invoking
    PowerShell. Failures are logged but suppressed so callers can continue with
    additional guardrails (e.g., ``--force``).
    @param description Human-readable text describing the restore point.
    @param dry_run Whether to simulate the request without executing PowerShell.
    @param timeout Maximum time to wait for PowerShell before aborting.
    @returns ``True`` if the restore point was created (or simulated), ``False``
    otherwise.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    description_text = description or "Office Janitor restore point"

    machine_logger.info(
        "restore_point.request",
        extra={
            "event": "restore_point.request",
            "description": description_text,
            "dry_run": dry_run,
        },
    )

    if dry_run:
        human_logger.info(
            "Dry-run enabled; would create system restore point: %s",
            description_text,
        )
        machine_logger.info(
            "restore_point.skipped",
            extra={
                "event": "restore_point.skipped",
                "reason": "dry_run",
                "description": description_text,
            },
        )
        return True

    if os.name != "nt":
        human_logger.info(
            "Restore points are only available on Windows hosts; skipping request.",
        )
        machine_logger.warning(
            "restore_point.unsupported_platform",
            extra={
                "event": "restore_point.unsupported_platform",
                "platform": os.name,
            },
        )
        return False

    command = _build_powershell_command(description_text)

    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=max(1, int(timeout)),
            check=False,
        )
    except FileNotFoundError:
        human_logger.warning(
            "%s unavailable; skipping system restore point creation.",
            _POWERSHELL_EXECUTABLE,
        )
        machine_logger.warning(
            "restore_point.unavailable",
            extra={
                "event": "restore_point.unavailable",
                "reason": "missing_powershell",
                "executable": _POWERSHELL_EXECUTABLE,
            },
        )
        return False
    except subprocess.TimeoutExpired:
        human_logger.warning(
            "System restore point creation timed out after %s seconds.", timeout
        )
        machine_logger.warning(
            "restore_point.timeout",
            extra={
                "event": "restore_point.timeout",
                "timeout": int(timeout),
            },
        )
        return False
    except Exception as exc:  # pragma: no cover - defensive fallback
        human_logger.warning("Unexpected restore point failure: %s", exc)
        machine_logger.warning(
            "restore_point.error",
            extra={
                "event": "restore_point.error",
                "error": repr(exc),
            },
        )
        return False

    stderr = (result.stderr or "").strip()
    stdout = (result.stdout or "").strip()

    if result.returncode == 0:
        human_logger.info("System restore point created: %s", description_text)
        machine_logger.info(
            "restore_point.success",
            extra={
                "event": "restore_point.success",
                "description": description_text,
                "stdout": stdout,
            },
        )
        return True

    human_logger.warning(
        "Restore point creation returned %s: %s",
        result.returncode,
        stderr or stdout or "(no error output)",
    )
    machine_logger.warning(
        "restore_point.failed",
        extra={
            "event": "restore_point.failed",
            "code": int(result.returncode),
            "stderr": stderr,
            "stdout": stdout,
        },
    )
    return False


def _build_powershell_command(description: str) -> Sequence[str]:
    """!
    @brief Construct the PowerShell command used to create a restore point.
    @details The generated script attempts the WMI ``SystemRestore`` path first
    to avoid policy blocks on ``Checkpoint-Computer``. If that fails the script
    falls back to ``Checkpoint-Computer`` before surfacing an error.
    @param description Description string for the restore point.
    @returns Sequence suitable for :func:`subprocess.run`.
    """

    description_literal = json.dumps(description)
    script = textwrap.dedent(
        f"""
        $description = {description_literal}
        $restorePointType = 0
        $eventType = 100
        try {{
            $systemRestore = Get-WmiObject -Class SystemRestore -Namespace \"root/default\" -ErrorAction Stop
            $result = $systemRestore.CreateRestorePoint($description, $restorePointType, $eventType)
            if ($result.ReturnValue -eq 0) {{
                exit 0
            }}
            exit $result.ReturnValue
        }} catch {{
            try {{
                Checkpoint-Computer -Description $description -RestorePointType 'MODIFY_SETTINGS' -ErrorAction Stop | Out-Null
                exit 0
            }} catch {{
                $message = $_.Exception.Message
                Write-Error $message
                exit 1
            }}
        }}
        """
    ).strip()

    return [
        _POWERSHELL_EXECUTABLE,
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        script,
    ]
