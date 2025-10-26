"""!
@brief Structured logging helpers for Office Janitor.
@details Implements the dual-stream pipeline described in :mod:`spec.md` using
rotating file handlers for human-readable text and JSONL telemetry output.
Startup metadata sourced from :mod:`office_janitor.version` is recorded so
automation can correlate log bundles.
"""
from __future__ import annotations

import datetime as _dt
import json
import logging
import sys
import uuid
from logging import handlers
from pathlib import Path
from typing import Dict, Iterable, Mapping, Tuple

from . import version

HUMAN_LOGGER_NAME = "office_janitor.human"
"""!
@brief Logger name for human-readable output.
"""

MACHINE_LOGGER_NAME = "office_janitor.machine"
"""!
@brief Logger name for JSONL telemetry output.
"""

_STANDARD_RECORD_KEYS: Dict[str, None] = {
    "name": None,
    "msg": None,
    "args": None,
    "levelname": None,
    "levelno": None,
    "pathname": None,
    "filename": None,
    "module": None,
    "exc_info": None,
    "exc_text": None,
    "stack_info": None,
    "lineno": None,
    "funcName": None,
    "created": None,
    "msecs": None,
    "relativeCreated": None,
    "thread": None,
    "threadName": None,
    "processName": None,
    "process": None,
    "message": None,
    "asctime": None,
    "channel": None,
}

_CURRENT_LOG_DIRECTORY: Path | None = None
_RUN_METADATA: Dict[str, object] | None = None


class _ChannelFilter(logging.Filter):
    """!
    @brief Inject a fixed ``channel`` attribute on log records.
    @details The human and machine loggers rely on this filter so formatters can
    emit a stable metadata field identifying the stream without callers needing
    to provide ``extra`` parameters manually.
    """

    def __init__(self, channel: str) -> None:
        super().__init__()
        self._channel = channel

    def filter(self, record: logging.LogRecord) -> bool:
        record.channel = self._channel
        return True


class _JsonLineFormatter(logging.Formatter):
    """!
    @brief Format ``LogRecord`` instances as single-line JSON objects.
    @details The formatter extracts standard metadata (timestamp, level, logger,
    and message) and merges any custom ``extra`` attributes provided by callers.
    Values that are not JSON serializable are coerced to their ``repr``.
    """

    def format(self, record: logging.LogRecord) -> str:  # noqa: D401 - concise override
        moment = _dt.datetime.fromtimestamp(record.created, tz=_dt.timezone.utc)
        payload: Dict[str, object] = {
            "timestamp": moment.isoformat(timespec="milliseconds").replace("+00:00", "Z"),
            "level": record.levelname,
            "logger": record.name,
            "message": record.getMessage(),
            "channel": getattr(record, "channel", "machine"),
        }

        extras = _extract_extras(record)
        payload.update(extras)
        try:
            return json.dumps(payload, ensure_ascii=False)
        except TypeError:
            sanitized = {key: _coerce_json(value) for key, value in payload.items()}
            return json.dumps(sanitized, ensure_ascii=False)


def _extract_extras(record: logging.LogRecord) -> Dict[str, object]:
    """!
    @brief Collect non-standard attributes from a log record.
    """

    extras: Dict[str, object] = {}
    for key, value in record.__dict__.items():
        if key in _STANDARD_RECORD_KEYS:
            continue
        extras[key] = value
    return extras


def _coerce_json(value: object) -> object:
    """!
    @brief Coerce arbitrary objects into a JSON-safe representation.
    """

    try:
        json.dumps(value)
    except TypeError:
        return repr(value)
    return value


def _configure_logger(logger: logging.Logger, formatter: logging.Formatter, handlers_to_add: Iterable[logging.Handler]) -> None:
    """!
    @brief Reset a logger and attach the supplied handlers.
    """

    for handler in list(logger.handlers):
        logger.removeHandler(handler)
        handler.close()
    for flt in list(logger.filters):
        logger.removeFilter(flt)
    for handler in handlers_to_add:
        handler.setFormatter(formatter)
        logger.addHandler(handler)
    logger.propagate = False


def setup_logging(
    root_dir: Path,
    *,
    json_to_stdout: bool = False,
    level: int = logging.INFO,
) -> Tuple[logging.Logger, logging.Logger]:
    """!
    @brief Set up human and machine loggers.
    @details The function returns a pair of ``logging.Logger`` objects for the
    human-readable and structured event channels respectively. The directory is
    created if it does not exist and rotated files are configured for both
    streams.
    """

    global _CURRENT_LOG_DIRECTORY

    root_dir.mkdir(parents=True, exist_ok=True)
    _CURRENT_LOG_DIRECTORY = root_dir

    human_logger = logging.getLogger(HUMAN_LOGGER_NAME)
    machine_logger = logging.getLogger(MACHINE_LOGGER_NAME)

    human_logger.setLevel(level)
    machine_logger.setLevel(level)

    human_formatter = logging.Formatter(
        "%(asctime)s %(levelname)-8s [%(channel)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    machine_formatter = _JsonLineFormatter()

    human_file = handlers.RotatingFileHandler(
        root_dir / "office-janitor.log",
        maxBytes=1_048_576,
        backupCount=5,
        encoding="utf-8",
    )
    machine_file = handlers.RotatingFileHandler(
        root_dir / "office-janitor.jsonl",
        maxBytes=1_048_576,
        backupCount=5,
        encoding="utf-8",
    )

    machine_handlers: list[logging.Handler] = [machine_file]
    if json_to_stdout:
        machine_handlers.append(logging.StreamHandler(stream=sys.stdout))

    _configure_logger(human_logger, human_formatter, [human_file])
    _configure_logger(machine_logger, machine_formatter, machine_handlers)

    human_logger.addFilter(_ChannelFilter("human"))
    machine_logger.addFilter(_ChannelFilter("machine"))

    _emit_run_metadata(human_logger, machine_logger)

    return human_logger, machine_logger


def get_human_logger() -> logging.Logger:
    """!
    @brief Retrieve the configured human-readable logger.
    @details The helper ensures downstream modules share the same formatting and
    metadata expectations established during :func:`setup_logging`.
    """

    return logging.getLogger(HUMAN_LOGGER_NAME)


def get_machine_logger() -> logging.Logger:
    """!
    @brief Retrieve the configured machine/JSON logger.
    @details Downstream callers can use this helper rather than creating ad-hoc
    structured logging configuration.
    """

    return logging.getLogger(MACHINE_LOGGER_NAME)


def get_log_directory() -> Path | None:
    """!
    @brief Return the most recently configured log directory, if any.
    """

    return _CURRENT_LOG_DIRECTORY


def get_run_metadata() -> Mapping[str, object] | None:
    """!
    @brief Return the most recent run metadata payload.
    @details The structure contains ``run_id`` (UUID4), ``timestamp`` in ISO-8601
    UTC form, and version/build identifiers sourced from
    :mod:`office_janitor.version`.
    """

    return dict(_RUN_METADATA) if _RUN_METADATA is not None else None


def _emit_run_metadata(human_logger: logging.Logger, machine_logger: logging.Logger) -> None:
    """!
    @brief Emit startup metadata to the configured loggers.
    @details A unique run identifier is generated so downstream tooling can
    correlate log streams. Metadata is persisted globally for later retrieval.
    """

    global _RUN_METADATA

    moment = _dt.datetime.now(tz=_dt.timezone.utc)
    _RUN_METADATA = {
        "run_id": uuid.uuid4().hex,
        "timestamp": moment.isoformat(timespec="milliseconds").replace("+00:00", "Z"),
        "version": version.__version__,
        "build": version.__build__,
        "python": sys.version.split()[0],
        "logdir": str(_CURRENT_LOG_DIRECTORY) if _CURRENT_LOG_DIRECTORY else None,
    }

    human_logger.info(
        "Office Janitor %s (%s) starting — run %s",
        version.__version__,
        version.__build__,
        _RUN_METADATA["run_id"],
    )
    if _CURRENT_LOG_DIRECTORY is not None:
        human_logger.info("Logs directory: %s", _CURRENT_LOG_DIRECTORY)

    machine_logger.info("run_start", extra={"event": "run_start", "run": dict(_RUN_METADATA)})

