"""!
@brief Structured logging helpers for Office Janitor.
@details Implements the dual-stream pipeline described in :mod:`spec.md` using
rotating file handlers for human-readable text and JSONL telemetry output.
Startup metadata sourced from :mod:`office_janitor.version` is recorded so
automation can correlate log bundles.
"""

from __future__ import annotations

import datetime as _dt
import getpass
import json
import logging
import os
import sys
import uuid
from collections import deque
from collections.abc import Iterable, Mapping, MutableMapping, MutableSequence
from logging import handlers
from pathlib import Path
from typing import (
    Callable,
)

from . import version

HUMAN_LOGGER_NAME = "office_janitor.human"
"""!
@brief Logger name for human-readable output.
"""

MACHINE_LOGGER_NAME = "office_janitor.machine"
"""!
@brief Logger name for JSONL telemetry output.
"""

_STANDARD_RECORD_KEYS: dict[str, None] = {
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
_RUN_METADATA: dict[str, object] | None = None
_SESSION_ID: str | None = None
_MACHINE_INFO: dict[str, str] | None = None
_UI_EVENT_EMITTER: Callable[..., object] | None = None
_UI_EVENT_QUEUE: MutableSequence[dict[str, object]] | deque[dict[str, object]] | None = None


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
        payload: dict[str, object] = {
            "timestamp": moment.isoformat(timespec="milliseconds").replace("+00:00", "Z"),
            "level": record.levelname,
            "logger": record.name,
            "message": record.getMessage(),
            "channel": getattr(record, "channel", "machine"),
        }

        if _MACHINE_INFO is not None:
            payload.setdefault("machine", dict(_MACHINE_INFO))
        if _SESSION_ID is not None:
            payload.setdefault("corr", _SESSION_ID)

        extras = _extract_extras(record)
        if _SESSION_ID is not None:
            session_info = extras.get("session")
            if not isinstance(session_info, Mapping):
                session_info = {}
            else:
                session_info = dict(session_info)
            session_info.setdefault("id", _SESSION_ID)
            extras["session"] = session_info
        payload.update(extras)
        try:
            return json.dumps(payload, ensure_ascii=False)
        except TypeError:
            sanitized = {key: _coerce_json(value) for key, value in payload.items()}
            return json.dumps(sanitized, ensure_ascii=False)


def _extract_extras(record: logging.LogRecord) -> dict[str, object]:
    """!
    @brief Collect non-standard attributes from a log record.
    """

    extras: dict[str, object] = {}
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


def _compute_machine_info() -> dict[str, str]:
    """!
    @brief Capture host and user metadata for log enrichment.
    """

    try:
        import platform

        platform_host = platform.node()
    except Exception:
        platform_host = ""

    host = os.environ.get("COMPUTERNAME") or os.environ.get("HOSTNAME") or platform_host
    try:
        user = getpass.getuser()
    except Exception:
        user = os.environ.get("USERNAME") or os.environ.get("USER") or ""
    info: dict[str, str] = {}
    if host:
        info["host"] = host
    if user:
        info["user"] = user
    return info


class _SizedTimedRotatingFileHandler(handlers.TimedRotatingFileHandler):
    """!
    @brief Combine size and time-based log rotation.
    @details The standard library does not ship a handler that supports both
    ``maxBytes`` and ``when`` rotation triggers simultaneously. This helper
    subclasses :class:`logging.handlers.TimedRotatingFileHandler` and extends
    :func:`shouldRollover` with the size check from
    :class:`logging.handlers.RotatingFileHandler`.
    """

    def __init__(
        self,
        filename: str | os.PathLike[str],
        *,
        max_bytes: int = 0,
        backup_count: int = 0,
        **kwargs: object,
    ) -> None:
        when = kwargs.pop("when", "midnight")
        interval = kwargs.pop("interval", 1)
        super().__init__(
            filename,
            when=when,
            interval=interval,
            backupCount=backup_count,
            encoding="utf-8",
            **kwargs,
        )
        self.maxBytes = max_bytes

    def shouldRollover(self, record: logging.LogRecord) -> int:  # noqa: D401 - stdlib compatibility
        if super().shouldRollover(record):
            return 1
        if self.maxBytes > 0:
            if self.stream is None:
                self.stream = self._open()
            msg = f"{self.format(record)}\n"
            self.stream.seek(0, os.SEEK_END)
            if self.stream.tell() + len(msg) >= self.maxBytes:
                return 1
        return 0


def _configure_logger(
    logger: logging.Logger, formatter: logging.Formatter, handlers_to_add: Iterable[logging.Handler]
) -> None:
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
) -> tuple[logging.Logger, logging.Logger]:
    """!
    @brief Set up human and machine loggers.
    @details The function returns a pair of ``logging.Logger`` objects for the
    human-readable and structured event channels respectively. The directory is
    created if it does not exist and rotated files are configured for both
    streams.
    """

    global _CURRENT_LOG_DIRECTORY, _MACHINE_INFO

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

    human_file = _SizedTimedRotatingFileHandler(
        root_dir / "human.log",
        max_bytes=10_485_760,
        backup_count=10,
    )
    machine_file = _SizedTimedRotatingFileHandler(
        root_dir / "events.jsonl",
        max_bytes=10_485_760,
        backup_count=10,
    )

    machine_handlers: list[logging.Handler] = [machine_file]
    if json_to_stdout:
        machine_handlers.append(logging.StreamHandler(stream=sys.stdout))

    _configure_logger(human_logger, human_formatter, [human_file])
    _configure_logger(machine_logger, machine_formatter, machine_handlers)

    human_logger.addFilter(_ChannelFilter("human"))
    machine_logger.addFilter(_ChannelFilter("machine"))

    global _SESSION_ID
    _SESSION_ID = uuid.uuid4().hex
    if _MACHINE_INFO is None:
        _MACHINE_INFO = _compute_machine_info()

    _emit_run_metadata(human_logger, machine_logger)

    return human_logger, machine_logger


def get_loggers(json_stdout: bool, level: int) -> tuple[logging.Logger, logging.Logger]:
    """!
    @brief Retrieve configured loggers, provisioning defaults when necessary.
    @details Modules may request stdout mirroring for JSONL events dynamically.
    If logging has not yet been configured, a ``logs`` directory under the
    current working directory is used as the fallback root.
    """

    default_dir = _CURRENT_LOG_DIRECTORY or (Path.cwd() / "logs")
    human = logging.getLogger(HUMAN_LOGGER_NAME)
    machine = logging.getLogger(MACHINE_LOGGER_NAME)

    if not human.handlers or not machine.handlers:
        setup_logging(default_dir, json_to_stdout=json_stdout, level=level)
        human = logging.getLogger(HUMAN_LOGGER_NAME)
        machine = logging.getLogger(MACHINE_LOGGER_NAME)
    else:
        human.setLevel(level)
        machine.setLevel(level)
        if json_stdout:
            _ensure_stdout_handler(machine)
        else:
            _remove_stdout_handler(machine)

    return human, machine


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


def register_ui_event_sink(
    *,
    emitter: Callable[..., object] | None = None,
    queue: MutableSequence[dict[str, object]] | deque[dict[str, object]] | None = None,
) -> None:
    """!
    @brief Register callables used to relay UI events to interactive surfaces.
    @details Modules may emit human-centric notifications via
    :func:`emit_ui_event`. When an ``emitter`` is provided, it is invoked first
    so higher level interfaces can react immediately. The optional ``queue``
    parameter captures events for later consumption by the CLI/TUI when no
    direct emitter exists. Passing no arguments resets previously registered
    sinks.
    @param emitter Callable mirroring ``emit_event`` in :mod:`office_janitor.main`.
    @param queue Mutable sequence that accepts dictionaries with ``event`` and
    ``message`` keys.
    """

    global _UI_EVENT_EMITTER, _UI_EVENT_QUEUE

    _UI_EVENT_EMITTER = emitter
    _UI_EVENT_QUEUE = queue


def emit_ui_event(event: str, message: str, **payload: object) -> bool:
    """!
    @brief Emit an event for consumption by CLI/TUI layers when available.
    @details The helper prefers an explicit emitter registered via
    :func:`register_ui_event_sink` and falls back to appending into the
    configured queue. A boolean return value indicates whether any sink accepted
    the event, allowing callers to decide if additional fallbacks are required.
    @param event Event identifier (for example ``"processes.outlook_reassurance"``).
    @param message Human-readable message describing the event.
    @returns ``True`` when the event was delivered to an emitter or queue.
    """

    delivered = False

    emitter = _UI_EVENT_EMITTER
    if callable(emitter):
        try:
            emitter(event, message=message, **payload)
            delivered = True
        except Exception:
            delivered = False

    queue = _UI_EVENT_QUEUE
    if queue is not None and hasattr(queue, "append"):
        record: dict[str, object] = {"event": event, "message": message}
        if payload:
            record["data"] = dict(payload)
        try:
            queue.append(record)  # type: ignore[arg-type]
            delivered = True
        except Exception:
            pass

    return delivered


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
    session_id = _SESSION_ID or uuid.uuid4().hex
    machine = _MACHINE_INFO or _compute_machine_info()
    _RUN_METADATA = {
        "session_id": session_id,
        "run_id": session_id,
        "timestamp": moment.isoformat(timespec="milliseconds").replace("+00:00", "Z"),
        "version": version.__version__,
        "build": version.__build__,
        "python": sys.version.split()[0],
        "logdir": str(_CURRENT_LOG_DIRECTORY) if _CURRENT_LOG_DIRECTORY else None,
        "machine": dict(machine),
    }

    human_logger.info(
        "Office Janitor %s (%s) starting — session %s",
        version.__version__,
        version.__build__,
        session_id,
    )
    if _CURRENT_LOG_DIRECTORY is not None:
        human_logger.info("Logs directory: %s", _CURRENT_LOG_DIRECTORY)

    machine_logger.info(
        "run_start",
        extra={
            "event": "run_start",
            "run": dict(_RUN_METADATA),
            "session": {"id": session_id},
            "machine": dict(machine),
        },
    )


def _ensure_stdout_handler(machine_logger: logging.Logger) -> None:
    """!
    @brief Guarantee a stdout stream handler using the JSON formatter.
    """

    for handler in machine_logger.handlers:
        if (
            isinstance(handler, logging.StreamHandler)
            and getattr(handler, "stream", None) is sys.stdout
        ):
            return
    stdout_handler = logging.StreamHandler(stream=sys.stdout)
    stdout_handler.setFormatter(_JsonLineFormatter())
    machine_logger.addHandler(stdout_handler)


def _remove_stdout_handler(machine_logger: logging.Logger) -> None:
    """!
    @brief Remove stdout stream handlers to avoid duplicate emission.
    """

    for handler in list(machine_logger.handlers):
        if (
            isinstance(handler, logging.StreamHandler)
            and getattr(handler, "stream", None) is sys.stdout
        ):
            machine_logger.removeHandler(handler)
            handler.close()


def build_event_extra(
    event: str,
    *,
    step_id: str | None = None,
    correlation: Mapping[str, object] | None = None,
    extra: Mapping[str, object] | None = None,
) -> dict[str, object]:
    """!
    @brief Compose ``extra`` payloads with consistent contextual metadata.
    @details The helper is intended for machine loggers so downstream telemetry
    consumers receive uniform ``event`` identifiers, optional ``step_id``, and
    correlation dictionaries. Run/session identifiers are injected
    automatically when available.
    """

    payload: dict[str, object] = {"event": event}
    if step_id is not None:
        payload["step_id"] = step_id
    if correlation:
        payload["correlation"] = dict(correlation)
    if extra:
        for key, value in extra.items():
            payload[key] = value

    if _RUN_METADATA:
        run_field = _merge_mapping(
            payload.get("run"),
            {
                "run_id": _RUN_METADATA.get("run_id"),
                "session_id": _RUN_METADATA.get("session_id"),
                "timestamp": _RUN_METADATA.get("timestamp"),
            },
        )
        payload["run"] = run_field
        session_field = _merge_mapping(
            payload.get("session"), {"id": _RUN_METADATA.get("session_id")}
        )
        payload["session"] = session_field

    return payload


def _merge_mapping(existing: object, defaults: Mapping[str, object | None]) -> dict[str, object]:
    """!
    @brief Merge mapping ``defaults`` into ``existing`` with fallbacks.
    """

    result: dict[str, object] = {}
    if isinstance(existing, MutableMapping):
        result.update(existing)
    for key, value in defaults.items():
        if value is None:
            continue
        result.setdefault(key, value)
    return result
