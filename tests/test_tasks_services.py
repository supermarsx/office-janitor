"""!
@brief Targeted tests for the tasks/services helpers.
@details Focus on timeout escalation behaviour when stopping services so the
reboot recommendation plumbing feeding the scrub summary remains covered.
"""

from __future__ import annotations

import json
import pathlib
import sys
from collections.abc import Sequence

PROJECT_ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import exec_utils, logging_ext, tasks_services  # noqa: E402


def _command_result(
    command: Sequence[str],
    *,
    returncode: int = 0,
    stdout: str = "",
    stderr: str = "",
    skipped: bool = False,
    timed_out: bool = False,
    error: str | None = None,
) -> exec_utils.CommandResult:
    """!
    @brief Helper to fabricate :class:`CommandResult` instances for tests.
    """

    return exec_utils.CommandResult(
        command=[str(part) for part in command],
        returncode=returncode,
        stdout=stdout,
        stderr=stderr,
        duration=0.0,
        skipped=skipped,
        timed_out=timed_out,
        error=error,
    )


def test_stop_services_timeout_requests_reboot(monkeypatch, tmp_path) -> None:
    """!
    @brief Timeouts should mark services for reboot and log the escalation.
    """

    logging_ext.setup_logging(tmp_path)
    commands: list[list[str]] = []

    def fake_run(command, *, event, **kwargs):
        commands.append([str(part) for part in command])
        if command[1] == "stop":
            return _command_result(
                command,
                timed_out=True,
                error="timeout",
            )
        return _command_result(command)

    monkeypatch.setattr(tasks_services.exec_utils, "run_command", fake_run)
    outcome = tasks_services.stop_services(["ClickToRunSvc"], timeout=5)

    assert commands[0][:2] == ["sc.exe", "stop"]
    assert commands[1][:2] == ["sc.exe", "config"]
    assert outcome == {
        "reboot_required": True,
        "services_requiring_reboot": ["ClickToRunSvc"],
    }

    human_log = (tmp_path / "human.log").read_text(encoding="utf-8")
    assert "recommend reboot" in human_log.lower()

    machine_log = (tmp_path / "events.jsonl").read_text(encoding="utf-8").splitlines()
    machine_events = [json.loads(line) for line in machine_log if line.strip()]
    timeout_event = next(
        event for event in machine_events if event.get("event") == "service_stop_timeout"
    )
    assert timeout_event["reboot_required"] is True
    assert timeout_event["service"] == "ClickToRunSvc"

    # Ensure the global accumulator reports the service and is cleared for later tests.
    assert tasks_services.consume_reboot_recommendations() == ["ClickToRunSvc"]
    assert tasks_services.consume_reboot_recommendations() == []
