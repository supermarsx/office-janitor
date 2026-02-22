"""!
@brief Cross-locale encoding helpers for subprocess output.
@details Windows command-line utilities may emit output in the system OEM
codepage, the ANSI codepage, or UTF-8 depending on locale and application
settings.  When the active codepage cannot represent every byte in the
stream (e.g. ``cp1252`` encountering ``0x81`` on German systems), the
default ``text=True`` behaviour of :mod:`subprocess` raises a
``UnicodeDecodeError``.

This module provides constants and helpers that ensure subprocess output is
decoded without data-loss crashes.  The recommended approach is to use
explicit ``encoding`` / ``errors`` parameters on every :class:`subprocess.Popen`
or :func:`subprocess.run` call that captures output.
"""

from __future__ import annotations

import locale
import logging
import sys

__all__ = [
    "SUBPROCESS_ENCODING",
    "SUBPROCESS_ERRORS",
    "safe_decode",
    "safe_subprocess_args",
    "get_system_encoding",
]

_logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Encoding strategy
# ---------------------------------------------------------------------------
# Use UTF-8 as the primary encoding.  PowerShell 5.1+ on modern Windows
# emits UTF-8 when asked via ``-Command`` and ConvertTo-Json, while cmd.exe
# tools default to the OEM codepage.  UTF-8 will correctly handle both
# in the vast majority of cases.
#
# ``errors="replace"`` ensures that any byte sequence that is *not* valid
# UTF-8 (or whichever encoding is in effect) is replaced with U+FFFD rather
# than raising ``UnicodeDecodeError``.  This trades perfect fidelity for
# crash-free execution — an acceptable trade-off because office-janitor
# only parses structured fragments from stdout and does not need byte-exact
# reproduction of foreign-locale output.
# ---------------------------------------------------------------------------

SUBPROCESS_ENCODING: str = "utf-8"
"""Encoding passed to :class:`subprocess.Popen` / :func:`subprocess.run`."""

SUBPROCESS_ERRORS: str = "replace"
"""Error handler passed to :class:`subprocess.Popen` / :func:`subprocess.run`.
``'replace'`` swaps undecodable bytes with U+FFFD instead of raising."""


def get_system_encoding() -> str:
    """!
    @brief Return the system's preferred encoding, falling back to ``utf-8``.
    @details Useful for informational logging so operators can see which locale
    the host was configured for.
    """
    try:
        encoding = locale.getpreferredencoding(False) or "utf-8"
    except Exception:  # pragma: no cover — exotic locale misconfiguration
        encoding = "utf-8"
    return encoding


def safe_decode(data: bytes | str, *, encoding: str | None = None) -> str:
    """!
    @brief Decode *data* to ``str`` without raising on invalid sequences.
    @details If *data* is already a ``str`` it is returned unchanged.
    The function attempts the requested (or default) encoding first, then
    falls back through ``utf-8`` and ``latin-1`` (which never fails) to
    guarantee a usable string is always returned.

    @param data  Raw bytes or a string.
    @param encoding  First encoding to attempt.  Defaults to
        :data:`SUBPROCESS_ENCODING`.
    @returns A decoded string.  Undecodable bytes are replaced with U+FFFD.
    """
    if isinstance(data, str):
        return data

    if encoding is None:
        encoding = SUBPROCESS_ENCODING

    # Attempt chain: requested → utf-8 → latin-1 (infallible).
    for enc, err_mode in [
        (encoding, "replace"),
        ("utf-8", "replace"),
        ("latin-1", "replace"),
    ]:
        try:
            return data.decode(enc, errors=err_mode)
        except (LookupError, UnicodeDecodeError):
            continue

    # latin-1 with 'replace' literally cannot fail, but just in case:
    return data.decode("latin-1", errors="replace")  # pragma: no cover


def safe_subprocess_args(
    *,
    capture: bool = True,
    encoding: str | None = None,
    errors: str | None = None,
) -> dict[str, object]:
    """!
    @brief Build keyword arguments suitable for ``subprocess.Popen`` or ``subprocess.run``.
    @details Returns a dict containing the ``encoding`` and ``errors`` keys so
    callers can unpack them with ``**``.  When *capture* is ``True`` the dict
    also includes ``stdout`` and ``stderr`` set to ``subprocess.PIPE``.

    Usage::

        import subprocess
        from office_janitor.encoding_helpers import safe_subprocess_args

        proc = subprocess.run(
            ["tasklist"],
            **safe_subprocess_args(),
        )

    @param capture  Include ``stdout=PIPE, stderr=PIPE`` in the returned dict.
    @param encoding  Override encoding (default :data:`SUBPROCESS_ENCODING`).
    @param errors  Override error handler (default :data:`SUBPROCESS_ERRORS`).
    @returns Keyword-argument dict ready for ``subprocess.Popen`` / ``run``.
    """
    import subprocess as _subprocess  # local to avoid module-level circular

    args: dict[str, object] = {
        "encoding": encoding or SUBPROCESS_ENCODING,
        "errors": errors or SUBPROCESS_ERRORS,
    }
    if capture:
        args["stdout"] = _subprocess.PIPE
        args["stderr"] = _subprocess.PIPE
    return args


def log_encoding_info() -> None:
    """!
    @brief Emit a one-time debug message describing the host's locale encoding.
    @details Aids triage of encoding-related bugs on non-English Windows
    installations.
    """
    preferred = get_system_encoding()
    stdout_enc = getattr(sys.stdout, "encoding", "unknown")
    _logger.debug(
        "Host encoding info — preferred: %s, stdout: %s, target: %s",
        preferred,
        stdout_enc,
        SUBPROCESS_ENCODING,
    )
