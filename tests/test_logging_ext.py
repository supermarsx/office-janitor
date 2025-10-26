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
    machine_logger.info("startup", extra={"event": "startup", "data": {"mode": "auto"}})
    _flush(human_logger)
    _flush(machine_logger)

    human_log = tmp_path / "office-janitor.log"
    machine_log = tmp_path / "office-janitor.jsonl"

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
    assert startup_entry["data"] == {"mode": "auto"}

    metadata = logging_ext.get_run_metadata()
    assert metadata is not None
    assert metadata["run_id"] == run_entry["run"]["run_id"]
    assert metadata["version"]
    assert metadata["logdir"].endswith(str(tmp_path))

    assert any(
        isinstance(handler, logging.handlers.RotatingFileHandler) for handler in human_logger.handlers
    )
    assert any(
        isinstance(handler, logging.handlers.RotatingFileHandler) for handler in machine_logger.handlers
    )


def test_json_stdout_mirror(tmp_path) -> None:
    """!
    @brief Verify machine logs can be mirrored to stdout when requested.
    """

    buffer = io.StringIO()
    with redirect_stdout(buffer):
        _, machine_logger = logging_ext.setup_logging(tmp_path, json_to_stdout=True)
        machine_logger.warning("mirror", extra={"event": "mirror"})
        _flush(machine_logger)

    output_lines = [line for line in buffer.getvalue().splitlines() if line.strip()]
    assert output_lines, "Expected JSONL output on stdout"
    first = json.loads(output_lines[0])
    assert first["event"] == "run_start"
    parsed = json.loads(output_lines[-1])
    assert parsed["event"] == "mirror"
    assert parsed["channel"] == "machine"


def test_logger_helpers_return_configured_instances(tmp_path) -> None:
    """!
    @brief Helper accessors should mirror the configured loggers.
    """

    human_logger, machine_logger = logging_ext.setup_logging(tmp_path)
    assert logging_ext.get_human_logger() is human_logger
    assert logging_ext.get_machine_logger() is machine_logger
