"""!
@brief Tests for :mod:`office_janitor.logging_ext`.
"""
from __future__ import annotations

import io
import json
import logging
import pathlib
import sys
from contextlib import redirect_stdout

import pytest

PROJECT_ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import logging_ext


@pytest.fixture(autouse=True)
def _reset_logging_state() -> None:
    """!
    @brief Reset logging between tests to avoid handler leakage.
    """

    yield
    logging.shutdown()
    for name in (logging_ext.HUMAN_LOGGER_NAME, logging_ext.MACHINE_LOGGER_NAME):
        logger = logging.getLogger(name)
        for handler in list(logger.handlers):
            logger.removeHandler(handler)
            handler.close()
        for flt in list(logger.filters):
            logger.removeFilter(flt)


def _flush(logger: logging.Logger) -> None:
    """!
    @brief Flush all handlers for deterministic file writes.
    """

    for handler in logger.handlers:
        handler.flush()


def test_setup_logging_creates_files_and_formats(tmp_path) -> None:
    """!
    @brief Ensure setup creates log files and applies expected formatting.
    """

    human_logger, machine_logger = logging_ext.setup_logging(tmp_path)
    human_logger.info("hello world")
    machine_logger.info(
        "startup",
        extra=logging_ext.build_event_extra("startup", correlation={"mode": "auto"}),
    )
    _flush(human_logger)
    _flush(machine_logger)

    human_log = tmp_path / "human.log"
    machine_log = tmp_path / "events.jsonl"

    assert human_log.exists()
    assert machine_log.exists()

    human_text = human_log.read_text(encoding="utf-8")
    assert "hello world" in human_text
    assert "[human]" in human_text

    machine_entries = [
        json.loads(line)
        for line in machine_log.read_text(encoding="utf-8").splitlines()
        if line.strip()
    ]
    run_entry = machine_entries[0]
    assert run_entry["event"] == "run_start"
    assert "run" in run_entry
    assert run_entry["run"]["run_id"]

    startup_entry = next(item for item in machine_entries if item.get("event") == "startup")
    assert startup_entry["channel"] == "machine"
    assert startup_entry["correlation"] == {"mode": "auto"}
    assert startup_entry["session"]["id"] == run_entry["run"]["session_id"]

    metadata = logging_ext.get_run_metadata()
    assert metadata is not None
    assert metadata["run_id"] == run_entry["run"]["run_id"]
    assert metadata["session_id"] == run_entry["run"]["session_id"]
    assert metadata["version"]
    assert metadata["logdir"].endswith(str(tmp_path))

    assert any(
        isinstance(handler, logging.handlers.TimedRotatingFileHandler) for handler in human_logger.handlers
    )
    assert any(
        isinstance(handler, logging.handlers.TimedRotatingFileHandler) for handler in machine_logger.handlers
    )


def test_json_stdout_mirror(tmp_path) -> None:
    """!
    @brief Verify machine logs can be mirrored to stdout when requested.
    """

    buffer = io.StringIO()
    with redirect_stdout(buffer):
        _, machine_logger = logging_ext.setup_logging(tmp_path, json_to_stdout=True)
        machine_logger.warning("mirror", extra=logging_ext.build_event_extra("mirror"))
        _flush(machine_logger)

    output_lines = [line for line in buffer.getvalue().splitlines() if line.strip()]
    assert output_lines, "Expected JSONL output on stdout"
    first = json.loads(output_lines[0])
    assert first["event"] == "run_start"
    parsed = json.loads(output_lines[-1])
    assert parsed["event"] == "mirror"
    assert parsed["channel"] == "machine"
    assert parsed["session"]["id"] == first["run"]["session_id"]


def test_logger_helpers_return_configured_instances(tmp_path) -> None:
    """!
    @brief Helper accessors should mirror the configured loggers.
    """

    human_logger, machine_logger = logging_ext.setup_logging(tmp_path)
    assert logging_ext.get_human_logger() is human_logger
    assert logging_ext.get_machine_logger() is machine_logger


def test_build_event_extra_injects_run_context(tmp_path) -> None:
    """!
    @brief Ensure helper merges run/session metadata with caller supplied context.
    """

    logging_ext.setup_logging(tmp_path)
    payload = logging_ext.build_event_extra(
        "step_progress",
        step_id="registry-scan",
        correlation={"host": "test"},
        extra={"progress": 50},
    )

    assert payload["event"] == "step_progress"
    assert payload["step_id"] == "registry-scan"
    assert payload["correlation"] == {"host": "test"}
    assert payload["progress"] == 50
    assert "run" in payload
    assert payload["run"]["session_id"] == payload["session"]["id"]


def test_get_loggers_adds_stdout_when_requested(tmp_path) -> None:
    """!
    @brief Requesting JSON stdout mirroring attaches a stream handler on demand.
    """

    logging_ext.setup_logging(tmp_path, json_to_stdout=False)
    _, machine = logging_ext.get_loggers(json_stdout=True, level=logging.INFO)
    assert any(
        isinstance(handler, logging.StreamHandler) and getattr(handler, "stream", None) is sys.stdout
        for handler in machine.handlers
    )

    _, machine = logging_ext.get_loggers(json_stdout=False, level=logging.INFO)
    assert not any(
        isinstance(handler, logging.StreamHandler) and getattr(handler, "stream", None) is sys.stdout
        for handler in machine.handlers
    )
