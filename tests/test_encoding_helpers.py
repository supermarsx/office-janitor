"""!
@brief Tests for :mod:`office_janitor.encoding_helpers`.
@details Validates encoding constants, safe decode fallbacks, subprocess
argument builder, and system encoding discovery.
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import encoding_helpers  # noqa: E402

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------


class TestConstants:
    """!
    @brief Validate module-level encoding constants.
    """

    def test_subprocess_encoding_is_utf8(self) -> None:
        assert encoding_helpers.SUBPROCESS_ENCODING == "utf-8"

    def test_subprocess_errors_is_replace(self) -> None:
        assert encoding_helpers.SUBPROCESS_ERRORS == "replace"


# ---------------------------------------------------------------------------
# get_system_encoding
# ---------------------------------------------------------------------------


class TestGetSystemEncoding:
    """!
    @brief Validate system encoding discovery.
    """

    def test_returns_string(self) -> None:
        result = encoding_helpers.get_system_encoding()
        assert isinstance(result, str)
        assert len(result) > 0

    def test_fallback_on_failure(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """When locale.getpreferredencoding raises, fall back to utf-8."""
        import locale

        monkeypatch.setattr(locale, "getpreferredencoding", lambda _: None)
        assert encoding_helpers.get_system_encoding() == "utf-8"


# ---------------------------------------------------------------------------
# safe_decode
# ---------------------------------------------------------------------------


class TestSafeDecode:
    """!
    @brief Ensure safe_decode handles every edge case without raising.
    """

    def test_str_passthrough(self) -> None:
        """Strings are returned unchanged."""
        assert encoding_helpers.safe_decode("hello") == "hello"

    def test_valid_utf8_bytes(self) -> None:
        assert encoding_helpers.safe_decode(b"hello world") == "hello world"

    def test_german_umlauts_utf8(self) -> None:
        """German text encoded as UTF-8 must decode cleanly."""
        text = "Ärger mit Ü und Ö"
        raw = text.encode("utf-8")
        assert encoding_helpers.safe_decode(raw) == text

    def test_latin1_bytes_replaced_gracefully(self) -> None:
        """Bytes outside UTF-8 produce replacements instead of errors."""
        # 0x81 is valid latin-1 but not valid UTF-8; should not crash
        raw = b"hello\x81world"
        result = encoding_helpers.safe_decode(raw)
        assert "hello" in result
        assert "world" in result
        # Must not raise

    def test_cp1252_problematic_byte_0x81(self) -> None:
        """The exact byte that triggered the original German-locale bug."""
        # Simulate a chunk of output with 0x81 embedded
        raw = b"Scanning Click-to-Run installations\x81 done"
        result = encoding_helpers.safe_decode(raw)
        assert "Scanning" in result
        assert "done" in result

    def test_empty_bytes(self) -> None:
        assert encoding_helpers.safe_decode(b"") == ""

    def test_empty_string(self) -> None:
        assert encoding_helpers.safe_decode("") == ""

    def test_mixed_high_bytes(self) -> None:
        """Various high bytes that differ between cp1252 and utf-8."""
        raw = bytes(range(128, 256))
        result = encoding_helpers.safe_decode(raw)
        assert isinstance(result, str)
        # Must not raise

    def test_explicit_encoding_override(self) -> None:
        """Caller can specify a different encoding."""
        text = "café"
        raw = text.encode("latin-1")
        result = encoding_helpers.safe_decode(raw, encoding="latin-1")
        assert result == text

    def test_null_bytes(self) -> None:
        """Embedded NUL bytes should not crash."""
        raw = b"a\x00b\x00c"
        result = encoding_helpers.safe_decode(raw)
        assert "a" in result
        assert "c" in result


# ---------------------------------------------------------------------------
# safe_subprocess_args
# ---------------------------------------------------------------------------


class TestSafeSubprocessArgs:
    """!
    @brief Validate the kwarg builder for subprocess calls.
    """

    def test_capture_mode(self) -> None:
        args = encoding_helpers.safe_subprocess_args(capture=True)
        assert args["encoding"] == "utf-8"
        assert args["errors"] == "replace"
        assert args["stdout"] == subprocess.PIPE
        assert args["stderr"] == subprocess.PIPE

    def test_no_capture_mode(self) -> None:
        args = encoding_helpers.safe_subprocess_args(capture=False)
        assert args["encoding"] == "utf-8"
        assert args["errors"] == "replace"
        assert "stdout" not in args
        assert "stderr" not in args

    def test_custom_encoding(self) -> None:
        args = encoding_helpers.safe_subprocess_args(encoding="latin-1", errors="ignore")
        assert args["encoding"] == "latin-1"
        assert args["errors"] == "ignore"


# ---------------------------------------------------------------------------
# log_encoding_info (smoke test)
# ---------------------------------------------------------------------------


class TestLogEncodingInfo:
    """!
    @brief Verify log_encoding_info emits without crashing.
    """

    def test_does_not_raise(self) -> None:
        encoding_helpers.log_encoding_info()


# ---------------------------------------------------------------------------
# Integration: exec_utils uses encoding params
# ---------------------------------------------------------------------------


class TestExecUtilsEncodingIntegration:
    """!
    @brief Verify that exec_utils passes encoding params to Popen.
    """

    def test_popen_receives_encoding_and_errors(self, monkeypatch: pytest.MonkeyPatch) -> None:
        from office_janitor import exec_utils

        captured_kwargs: dict[str, object] = {}

        class CapturingPopen:
            def __init__(self, command, **kwargs):  # type: ignore[no-untyped-def]
                captured_kwargs.update(kwargs)
                self.returncode = 0

            def communicate(self, timeout=None):  # type: ignore[no-untyped-def]
                return ("ok", "")

        class _StubLogger:
            def info(self, *a, **kw):  # type: ignore[no-untyped-def]
                pass

            def warning(self, *a, **kw):  # type: ignore[no-untyped-def]
                pass

            def error(self, *a, **kw):  # type: ignore[no-untyped-def]
                pass

        monkeypatch.setattr(exec_utils.logging_ext, "get_human_logger", lambda: _StubLogger())
        monkeypatch.setattr(exec_utils.logging_ext, "get_machine_logger", lambda: _StubLogger())
        monkeypatch.setattr(exec_utils.subprocess, "Popen", CapturingPopen)

        exec_utils.run_command(["echo", "test"], event="enc_test")

        assert captured_kwargs.get("encoding") == "utf-8"
        assert captured_kwargs.get("errors") == "replace"
        # text=True should NOT be present (replaced by explicit encoding)
        assert "text" not in captured_kwargs

    def test_problematic_bytes_do_not_crash(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """Simulate output containing byte 0x81 that crashed on German Windows."""
        from office_janitor import exec_utils

        class ProblematicPopen:
            def __init__(self, command, **kwargs):  # type: ignore[no-untyped-def]
                self.returncode = 0
                # With encoding='utf-8' and errors='replace', the Popen object
                # would have already decoded the output. Simulate the decoded
                # result containing a replacement character.
                self._stdout = "Scanning\ufffdresult"
                self._stderr = ""

            def communicate(self, timeout=None):  # type: ignore[no-untyped-def]
                return (self._stdout, self._stderr)

        class _StubLogger:
            def info(self, *a, **kw):  # type: ignore[no-untyped-def]
                pass

            def warning(self, *a, **kw):  # type: ignore[no-untyped-def]
                pass

            def error(self, *a, **kw):  # type: ignore[no-untyped-def]
                pass

        monkeypatch.setattr(exec_utils.logging_ext, "get_human_logger", lambda: _StubLogger())
        monkeypatch.setattr(exec_utils.logging_ext, "get_machine_logger", lambda: _StubLogger())
        monkeypatch.setattr(exec_utils.subprocess, "Popen", ProblematicPopen)

        # This must not raise UnicodeDecodeError
        result = exec_utils.run_command(["some", "command"], event="german_test")
        assert result.returncode == 0
        assert "Scanning" in result.stdout


# ---------------------------------------------------------------------------
# Locale-specific byte sequences
# ---------------------------------------------------------------------------


class TestLocaleSpecificDecoding:
    """!
    @brief Validate safe_decode against byte sequences that appear in real
    Windows command output across different locale codepages.
    """

    def test_german_cp850_oem_output(self) -> None:
        """German OEM codepage (cp850) output from cmd.exe tools."""
        # cp850 encodes ä as 0x84, ö as 0x94, ü as 0x81, ß as 0xe1
        raw = b"Aufgabenplanung: 3 Auftr\x84ge ge\x84ndert, L\x94sung verf\x81gbar"
        result = encoding_helpers.safe_decode(raw)
        assert isinstance(result, str)
        assert "Aufgabenplanung" in result

    def test_german_cp1252_ansi_output(self) -> None:
        """German ANSI codepage (cp1252) output — 0x81 is undefined there."""
        raw = b"Microsoft Office Professional Plus 2019\x81"
        result = encoding_helpers.safe_decode(raw)
        assert "Microsoft Office" in result

    def test_french_cp1252_accents(self) -> None:
        """French accented characters in cp1252."""
        raw = "Vérification terminée".encode("cp1252")
        result = encoding_helpers.safe_decode(raw)
        assert isinstance(result, str)
        # UTF-8 decode of cp1252 bytes will produce replacement chars but not crash
        assert len(result) > 0

    def test_japanese_shift_jis_bytes(self) -> None:
        """Shift-JIS bytes from Japanese Windows should not crash."""
        raw = b"\x83\x49\x83\x74\x83\x42\x83\x58"  # "オフィス" in Shift-JIS
        result = encoding_helpers.safe_decode(raw)
        assert isinstance(result, str)
        assert len(result) > 0

    def test_chinese_gbk_bytes(self) -> None:
        """GBK bytes from Chinese Windows should not crash."""
        raw = b"\xb0\xec\xb9\xab\xc8\xed\xbc\xfe"  # "办公软件" in GBK
        result = encoding_helpers.safe_decode(raw)
        assert isinstance(result, str)
        assert len(result) > 0

    def test_korean_euc_kr_bytes(self) -> None:
        """EUC-KR bytes from Korean Windows should not crash."""
        raw = b"\xbf\xc0\xc7\xc7\xbd\xba"  # "오피스" in EUC-KR
        result = encoding_helpers.safe_decode(raw)
        assert isinstance(result, str)

    def test_russian_cp1251_bytes(self) -> None:
        """cp1251 bytes from Russian Windows should not crash."""
        raw = b"\xce\xf4\xe8\xf1"  # "Офис" in cp1251
        result = encoding_helpers.safe_decode(raw)
        assert isinstance(result, str)

    def test_mixed_ascii_and_high_bytes(self) -> None:
        """Typical tool output: ASCII with occasional high bytes interspersed."""
        raw = (
            b"[    0.095124]       Checking Office services [  OK  ]\r\n"
            b"[    0.572452]       Scanning Click-to-Run \x81installations [  OK  ]\r\n"
            b"[    0.358554]       Scanning AppX/MSIX packages [  OK  ]\r\n"
        )
        result = encoding_helpers.safe_decode(raw)
        assert "Checking Office services" in result
        assert "Scanning Click-to-Run" in result
        assert "Scanning AppX/MSIX packages" in result

    def test_powershell_utf8_bom(self) -> None:
        """PowerShell sometimes emits a UTF-8 BOM prefix."""
        raw = b'\xef\xbb\xbf{"Name":"Office"}'
        result = encoding_helpers.safe_decode(raw)
        assert "Office" in result

    def test_windows_newlines_preserved(self) -> None:
        """CRLF line endings must survive decoding."""
        raw = b"line1\r\nline2\r\nline3\r\n"
        result = encoding_helpers.safe_decode(raw)
        assert "line1" in result
        assert "line2" in result
        assert "line3" in result

    def test_all_single_byte_values(self) -> None:
        """Every possible single-byte value (0x00–0xFF) must decode without error."""
        for byte_val in range(256):
            raw = bytes([byte_val])
            result = encoding_helpers.safe_decode(raw)
            assert isinstance(result, str)

    def test_truncated_multibyte_utf8(self) -> None:
        """Truncated multi-byte UTF-8 sequences should be replaced, not crash."""
        # Start of 3-byte sequence (0xE2) without continuation bytes
        raw = b"hello\xe2world"
        result = encoding_helpers.safe_decode(raw)
        assert "hello" in result
        assert "world" in result

    def test_overlong_utf8_sequence(self) -> None:
        """Overlong UTF-8 encoding should be handled gracefully."""
        # Overlong encoding of '/' (0x2F) as 2 bytes: 0xC0 0xAF
        raw = b"path\xc0\xaffile"
        result = encoding_helpers.safe_decode(raw)
        assert "path" in result
        assert "file" in result


# ---------------------------------------------------------------------------
# safe_decode with specific encoding override
# ---------------------------------------------------------------------------


class TestSafeDecodeEncodingFallback:
    """!
    @brief Test the fallback chain when the requested encoding is invalid.
    """

    def test_bogus_encoding_falls_back(self) -> None:
        """An unrecognised encoding name should not crash."""
        raw = b"hello"
        result = encoding_helpers.safe_decode(raw, encoding="totally-bogus-codec-12345")
        assert result == "hello"

    def test_latin1_never_fails(self) -> None:
        """latin-1 can decode any byte; verify the final fallback works."""
        raw = bytes(range(256))
        result = encoding_helpers.safe_decode(raw, encoding="latin-1")
        assert len(result) == 256


# ---------------------------------------------------------------------------
# Integration: appx_uninstall encoding
# ---------------------------------------------------------------------------


class TestAppxUninstallEncoding:
    """!
    @brief Verify appx_uninstall._run_powershell uses encoding params.
    """

    def test_run_powershell_uses_utf8_encoding(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """_run_powershell should pass encoding/errors instead of text=True."""
        from office_janitor import appx_uninstall

        captured_kwargs: dict[str, object] = {}

        def capturing_run(*args: object, **kwargs: object) -> subprocess.CompletedProcess[str]:
            captured_kwargs.update(kwargs)
            return subprocess.CompletedProcess(args=args, returncode=0, stdout="", stderr="")

        monkeypatch.setattr(subprocess, "run", capturing_run)

        appx_uninstall._run_powershell("Write-Host test")

        assert captured_kwargs.get("encoding") == "utf-8"
        assert captured_kwargs.get("errors") == "replace"
        assert "text" not in captured_kwargs

    def test_run_powershell_handles_high_bytes(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """PowerShell returning non-ASCII output should not crash."""
        from office_janitor import appx_uninstall

        def fake_run(*args: object, **kwargs: object) -> subprocess.CompletedProcess[str]:
            return subprocess.CompletedProcess(
                args=args,
                returncode=0,
                stdout='{"Name":"Microsoft.Office.Desktop\ufffd"}',
                stderr="",
            )

        monkeypatch.setattr(subprocess, "run", fake_run)

        result = appx_uninstall._run_powershell("Get-AppxPackage")
        assert result.returncode == 0
        assert "Microsoft.Office" in result.stdout


# ---------------------------------------------------------------------------
# Integration: odt_build encoding
# ---------------------------------------------------------------------------


class TestOdtBuildEncoding:
    """!
    @brief Verify odt_build subprocess calls use encoding params.
    """

    def test_tasklist_call_uses_encoding(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """_find_running_clicktorun_processes should use encoding/errors."""
        from office_janitor import odt_build

        captured_kwargs: dict[str, object] = {}

        def capturing_run(*args: object, **kwargs: object) -> subprocess.CompletedProcess[str]:
            captured_kwargs.update(kwargs)
            return subprocess.CompletedProcess(args=args, returncode=0, stdout="", stderr="")

        monkeypatch.setattr(subprocess, "run", capturing_run)

        odt_build._find_running_clicktorun_processes()

        assert captured_kwargs.get("encoding") == "utf-8"
        assert captured_kwargs.get("errors") == "replace"
        assert "text" not in captured_kwargs


# ---------------------------------------------------------------------------
# Large output simulation
# ---------------------------------------------------------------------------


class TestLargeOutputDecoding:
    """!
    @brief Stress-test safe_decode with large payloads containing scattered
    problematic bytes — simulating real scheduled-task or registry output.
    """

    def test_large_output_with_scattered_high_bytes(self) -> None:
        """4 KB of ASCII with random 0x81 bytes every ~100 chars."""
        chunks = []
        for _i in range(40):
            chunks.append(b"x" * 99 + b"\x81")
        raw = b"".join(chunks)
        assert len(raw) == 4000
        result = encoding_helpers.safe_decode(raw)
        assert isinstance(result, str)
        assert len(result) == 4000  # each byte maps to one char with replace

    def test_64kb_pure_ascii(self) -> None:
        """Large ASCII-only payload should decode perfectly."""
        raw = b"A" * 65536
        result = encoding_helpers.safe_decode(raw)
        assert result == "A" * 65536

    def test_registry_export_style_output(self) -> None:
        """Simulated registry export with paths containing high bytes."""
        lines = [
            b"HKLM\\SOFTWARE\\Microsoft\\Office\\16.0\\Common\\InstallRoot",
            b"    Path    REG_SZ    C:\\Program Files\\Microsoft Office\\root\\Office16\\",
            b"    Konfiguration\x81    REG_SZ    Standard",
            b"HKLM\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration",
            b"    Platform    REG_SZ    x64",
        ]
        raw = b"\r\n".join(lines)
        result = encoding_helpers.safe_decode(raw)
        assert "InstallRoot" in result
        assert "ClickToRun" in result
        assert "Platform" in result

    def test_scheduled_task_xml_with_encoding_issues(self) -> None:
        """Scheduled task output often contains locale-specific chars."""
        raw = (
            b'<?xml version="1.0" encoding="UTF-16"?>\r\n'
            b"<Task>\r\n"
            b"  <RegistrationInfo>\r\n"
            b"    <Description>Office Automatische Updates \x81berpr\x81fung</Description>\r\n"
            b"  </RegistrationInfo>\r\n"
            b"</Task>"
        )
        result = encoding_helpers.safe_decode(raw)
        assert "<Task>" in result
        assert "Office Automatische Updates" in result
        assert "</Task>" in result
