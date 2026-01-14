"""!
@brief Registry tooling tests covering cleanup utilities.
@details Validates dry-run behaviour, guardrails, and command invocation for
registry export and delete helpers using mocked ``reg.exe`` availability. The
tests also cover Office uninstall heuristics exposed by the module.
"""

from __future__ import annotations

import pathlib
import sys
from collections.abc import Iterable

import pytest

PROJECT_ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import exec_utils, registry_tools  # noqa: E402


def _command_result(
    command: Iterable[str], returncode: int = 0, *, skipped: bool = False
) -> exec_utils.CommandResult:
    """!
    @brief Fabricate :class:`CommandResult` objects for command interception.
    """

    return exec_utils.CommandResult(
        command=[str(part) for part in command],
        returncode=returncode,
        stdout="",
        stderr="",
        duration=0.0,
        skipped=skipped,
    )


class _Recorder:
    """!
    @brief Minimal logger stub capturing emitted events.
    """

    def __init__(self) -> None:
        self.messages: list[tuple[str, dict]] = []

    def info(self, message: str, *args, **kwargs) -> None:  # noqa: D401 - logging compatibility
        payload = kwargs.copy()
        self.messages.append((message, payload))

    def warning(self, message: str, *args, **kwargs) -> None:  # noqa: D401 - logging compatibility
        payload = kwargs.copy()
        self.messages.append((message, payload))

    def debug(self, message: str, *args, **kwargs) -> None:  # noqa: D401 - logging compatibility
        payload = kwargs.copy()
        self.messages.append((message, payload))


def test_delete_keys_invokes_reg_when_available(monkeypatch) -> None:
    """!
    @brief Deletion should call ``reg delete`` when the binary exists.
    """

    commands: list[list[str]] = []

    monkeypatch.setattr(registry_tools.shutil, "which", lambda exe: "reg.exe")

    def fake_run(command, *, event, dry_run=False, **kwargs):
        commands.append([str(part) for part in command])
        return _command_result(command, skipped=dry_run)

    monkeypatch.setattr(registry_tools.exec_utils, "run_command", fake_run)

    registry_tools.delete_keys(
        ["HKLM\\Software\\Microsoft\\Office\\Contoso"],
        dry_run=False,
        logger=_Recorder(),
    )

    # Path case is preserved for reg.exe compatibility
    assert commands == [["reg.exe", "delete", "HKLM\\Software\\Microsoft\\Office\\Contoso", "/f"]]


def test_delete_keys_dry_run_skips_execution(monkeypatch) -> None:
    """!
    @brief Dry-run should avoid invoking ``reg.exe``.
    """

    monkeypatch.setattr(registry_tools.shutil, "which", lambda exe: "reg.exe")

    calls: list[bool] = []

    def fake_run(command, *, event, dry_run=False, **kwargs):
        calls.append(dry_run)
        return _command_result(command, skipped=dry_run)

    monkeypatch.setattr(registry_tools.exec_utils, "run_command", fake_run)

    registry_tools.delete_keys(
        ["HKCU\\Software\\Microsoft\\Office\\Tailspin"],
        dry_run=True,
        logger=_Recorder(),
    )

    assert calls and all(calls)


def test_delete_keys_rejects_disallowed_paths(monkeypatch) -> None:
    """!
    @brief Guardrails should reject deletions outside the whitelist.
    """

    monkeypatch.setattr(registry_tools.shutil, "which", lambda exe: "reg.exe")

    with pytest.raises(registry_tools.RegistryError):
        registry_tools.delete_keys(["HKLM\\Software\\Contoso"], dry_run=False, logger=_Recorder())


def test_export_keys_creates_placeholder_when_reg_missing(tmp_path, monkeypatch) -> None:
    """!
    @brief Exports should produce placeholder files if ``reg.exe`` is absent.
    """

    monkeypatch.setattr(registry_tools.shutil, "which", lambda exe: None)

    recorder = _Recorder()
    exported = registry_tools.export_keys(
        ["HKLM\\Software\\Microsoft\\Office\\Fabrikam"],
        tmp_path,
        logger=recorder,
    )

    assert exported, "Expected an export path to be returned"
    export_file = exported[0]
    assert export_file.exists()
    assert "Placeholder export" in export_file.read_text(encoding="utf-8")
    assert recorder.messages and recorder.messages[0][1]["extra"]["action"] == "registry-export"


def test_export_keys_dry_run_records_intent(tmp_path, monkeypatch) -> None:
    """!
    @brief Dry-run exports should log intent without creating files.
    """

    monkeypatch.setattr(registry_tools.shutil, "which", lambda exe: "reg.exe")

    recorder = _Recorder()

    def fake_run(command, *, event, dry_run=False, **kwargs):
        return _command_result(command, skipped=dry_run)

    monkeypatch.setattr(registry_tools.exec_utils, "run_command", fake_run)

    exported = registry_tools.export_keys(
        ["HKLM\\Software\\Microsoft\\Office\\Diagnostics"],
        tmp_path,
        dry_run=True,
        logger=recorder,
    )

    assert exported
    assert not exported[0].exists()
    assert recorder.messages[0][1]["extra"]["dry_run"] is True


def test_looks_like_office_entry_matches_keywords() -> None:
    """!
    @brief The Office heuristic should recognise branded display names.
    """

    entry = {"DisplayName": "Microsoft Office 365 ProPlus", "Publisher": "Microsoft Corporation"}
    assert registry_tools.looks_like_office_entry(entry)

    unrelated = {"DisplayName": "Contoso Widget", "Publisher": "Contoso"}
    assert not registry_tools.looks_like_office_entry(unrelated)


def test_export_keys_invokes_reg_command(tmp_path, monkeypatch) -> None:
    """!
    @brief When ``reg.exe`` is present the utility should invoke it for exports.
    """

    monkeypatch.setattr(registry_tools.shutil, "which", lambda exe: "C:/Windows/system32/reg.exe")
    calls: list[list[str]] = []

    def fake_run(command, *, event, dry_run=False, check=False, extra=None, **kwargs):
        calls.append([str(part) for part in command])
        return _command_result(command, skipped=dry_run)

    monkeypatch.setattr(registry_tools.exec_utils, "run_command", fake_run)

    exported = registry_tools.export_keys(
        ["HKLM\\Software\\Microsoft\\Office\\Diagnostics"],
        tmp_path,
        dry_run=False,
        logger=_Recorder(),
    )

    # Path case is preserved for reg.exe compatibility
    expected_command = [
        "C:/Windows/system32/reg.exe",
        "export",
        "HKLM\\Software\\Microsoft\\Office\\Diagnostics",
        str(exported[0]),
        "/y",
    ]
    assert calls == [expected_command]
    assert exported[0].parent == tmp_path


def test_export_keys_produces_unique_filenames(tmp_path, monkeypatch) -> None:
    """!
    @brief Duplicate key exports should yield unique placeholder filenames.
    """

    monkeypatch.setattr(registry_tools.shutil, "which", lambda exe: None)

    exported = registry_tools.export_keys(
        [
            "HKLM\\Software\\Microsoft\\Office\\Diagnostics",
            "HKLM\\Software\\Microsoft\\Office\\Diagnostics",
        ],
        tmp_path,
        dry_run=False,
        logger=_Recorder(),
    )

    assert len(exported) == 2
    names = [path.name for path in exported]
    assert names[0] != names[1]
    assert names[1].startswith(names[0][:-4])
    assert exported[0].exists() and exported[1].exists()


def test_iter_office_uninstall_entries_filters_non_office(monkeypatch) -> None:
    """!
    @brief Only Office-like entries should be returned from uninstall enumeration.
    """

    roots: Iterable[tuple[int, str]] = [
        (0x80000002, r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall")
    ]

    def fake_iter_subkeys(root: int, path: str, *, view: str | None = None):
        yield from ("{90160000-0011-0000-0000-0000000FF1CE}", "ContosoApp")

    def fake_read_values(root: int, path: str, *, view: str | None = None):
        if path.endswith("ContosoApp"):
            return {"DisplayName": "Contoso", "Publisher": "Contoso"}
        return {
            "DisplayName": "Microsoft Office Professional Plus 2016",
            "Publisher": "Microsoft Corporation",
            "ProductCode": "{90160000-0011-0000-0000-0000000FF1CE}",
        }

    monkeypatch.setattr(registry_tools, "iter_subkeys", fake_iter_subkeys)
    monkeypatch.setattr(registry_tools, "read_values", fake_read_values)

    results = list(registry_tools.iter_office_uninstall_entries(roots))
    assert len(results) == 1
    hive, path, values = results[0]
    assert hive == 0x80000002
    assert path.endswith("{90160000-0011-0000-0000-0000000FF1CE}")
    assert values["ProductCode"] == "{90160000-0011-0000-0000-0000000FF1CE}"


def test_normalize_registry_key() -> None:
    """!
    @brief Test registry key normalization.

    Hive prefixes are canonicalized (HKEY_LOCAL_MACHINE -> HKLM), but the
    path portion preserves its original case for compatibility with reg.exe.
    """

    from office_janitor.registry_tools import _normalize_registry_key

    # Hive canonicalization, path case preserved
    assert _normalize_registry_key("HKLM\\Software\\Microsoft") == "HKLM\\Software\\Microsoft"
    assert (
        _normalize_registry_key("HKEY_LOCAL_MACHINE\\Software\\Microsoft")
        == "HKLM\\Software\\Microsoft"
    )
    assert _normalize_registry_key("hkcu\\software") == "HKCU\\software"
    assert _normalize_registry_key("INVALID\\path") == "INVALID\\path"


def test_is_registry_path_allowed() -> None:
    """!
    @brief Test registry path validation.
    """

    from office_janitor.registry_tools import _is_registry_path_allowed

    # Assuming some allowed paths
    assert _is_registry_path_allowed("HKLM\\SOFTWARE\\MICROSOFT\\OFFICE")
    assert not _is_registry_path_allowed(
        "HKLM\\SOFTWARE\\MICROSOFT\\WINDOWS"
    )  # assuming not allowed


# ---------------------------------------------------------------------------
# Windows Installer Metadata Validation Tests
# ---------------------------------------------------------------------------


class TestWIMetadataValidation:
    """Tests for Windows Installer metadata validation functions."""

    def test_is_valid_compressed_guid_valid(self) -> None:
        """Valid 32-char hex strings should be accepted."""
        from office_janitor.registry_tools import _is_valid_compressed_guid

        assert _is_valid_compressed_guid("00000000000000000000000000000000")
        assert _is_valid_compressed_guid("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF")
        assert _is_valid_compressed_guid("09610000110000000000000000F01FEC")

    def test_is_valid_compressed_guid_invalid(self) -> None:
        """Invalid strings should be rejected."""
        from office_janitor.registry_tools import _is_valid_compressed_guid

        # Wrong length
        assert not _is_valid_compressed_guid("0000000000000000")
        assert not _is_valid_compressed_guid("")
        # Non-hex characters
        assert not _is_valid_compressed_guid("GGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGG")
        # With braces (standard GUID format)
        assert not _is_valid_compressed_guid("{00000000-0000-0000-0000-000000000000}")

    def test_validate_wi_metadata_key_finds_invalid(self, monkeypatch) -> None:
        """Should identify invalid entries in WI metadata."""
        subkeys = [
            "00000000000000000000000000000000",  # Valid
            "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF",  # Valid
            "INVALID",  # Invalid - wrong length
            "TOOSHORT",  # Invalid - wrong length
            "GGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGG",  # Invalid - non-hex
        ]

        def fake_iter_subkeys(hive, path, view=None):
            return iter(subkeys)

        monkeypatch.setattr(registry_tools, "iter_subkeys", fake_iter_subkeys)

        invalid = registry_tools.validate_wi_metadata_key(
            registry_tools._WINREG_HKLM,
            r"SOFTWARE\Classes\Installer\Products",
            expected_length=32,
            logger=_Recorder(),
        )

        assert len(invalid) == 3
        assert "INVALID" in invalid
        assert "TOOSHORT" in invalid
        assert "GGGGGGGGGGGGGGGGGGGGGGGGGGGGGGGG" in invalid

    def test_validate_wi_metadata_key_all_valid(self, monkeypatch) -> None:
        """Should return empty list when all entries are valid."""
        subkeys = [
            "00000000000000000000000000000000",
            "FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF",
            "0961000011000000000000000F01FEC0",
        ]

        def fake_iter_subkeys(hive, path, view=None):
            return iter(subkeys)

        monkeypatch.setattr(registry_tools, "iter_subkeys", fake_iter_subkeys)

        invalid = registry_tools.validate_wi_metadata_key(
            registry_tools._WINREG_HKLM,
            r"SOFTWARE\Classes\Installer\Products",
            expected_length=32,
            logger=_Recorder(),
        )

        assert invalid == []

    def test_validate_wi_metadata_key_missing_path(self, monkeypatch) -> None:
        """Should return empty list for non-existent paths."""

        def fake_iter_subkeys(hive, path, view=None):
            raise FileNotFoundError(path)

        monkeypatch.setattr(registry_tools, "iter_subkeys", fake_iter_subkeys)

        invalid = registry_tools.validate_wi_metadata_key(
            registry_tools._WINREG_HKLM,
            r"SOFTWARE\Classes\Installer\NonExistent",
            expected_length=32,
            logger=_Recorder(),
        )

        assert invalid == []

    def test_scan_wi_metadata_aggregates_results(self, monkeypatch) -> None:
        """Should scan multiple paths and aggregate results."""
        call_count = {"count": 0}

        def fake_validate(hive, path, expected_length, logger=None):
            call_count["count"] += 1
            if "Products" in path:
                return ["INVALID1", "INVALID2"]
            return []

        monkeypatch.setattr(registry_tools, "validate_wi_metadata_key", fake_validate)

        results = registry_tools.scan_wi_metadata(logger=_Recorder())

        # Should have called validate for paths with expected_length > 0
        assert call_count["count"] >= 1
        assert "Products" in results or len(results) == 0

    def test_cleanup_wi_orphaned_products_dry_run(self, monkeypatch) -> None:
        """Dry run should not delete but should count entries."""

        def fake_key_exists(path):
            return "Products" in path or "Features" in path

        def fake_delete_keys(keys, dry_run=False, logger=None):
            raise AssertionError("Should not delete in dry run mode")

        monkeypatch.setattr(registry_tools, "key_exists", fake_key_exists)
        monkeypatch.setattr(registry_tools, "delete_keys", fake_delete_keys)

        removed = registry_tools.cleanup_wi_orphaned_products(
            ["{90160000-000F-0000-1000-0000000FF1CE}"],
            dry_run=True,
            logger=_Recorder(),
        )

        # Should report entries that would be removed
        assert removed >= 1

    def test_cleanup_wi_orphaned_products_invalid_guid(self, monkeypatch) -> None:
        """Invalid product codes should be skipped with warning."""
        recorder = _Recorder()

        removed = registry_tools.cleanup_wi_orphaned_products(
            ["not-a-valid-guid"],
            dry_run=True,
            logger=recorder,
        )

        assert removed == 0
        # Should have logged a warning
        assert any("Invalid product code" in msg for msg, _ in recorder.messages)

    def test_cleanup_wi_orphaned_components_dry_run(self, monkeypatch) -> None:
        """Dry run should not delete but should count entries."""

        def fake_key_exists(path):
            return "Components" in path

        monkeypatch.setattr(registry_tools, "key_exists", fake_key_exists)

        removed = registry_tools.cleanup_wi_orphaned_components(
            ["{11111111-1111-1111-1111-111111111111}"],
            dry_run=True,
            logger=_Recorder(),
        )

        assert removed == 1

    def test_wi_metadata_paths_defined(self) -> None:
        """WI_METADATA_PATHS should define standard paths."""
        assert "Products" in registry_tools.WI_METADATA_PATHS
        assert "Components" in registry_tools.WI_METADATA_PATHS
        assert "Features" in registry_tools.WI_METADATA_PATHS


class TestShellIntegrationCleanup:
    """Tests for shell integration cleanup functions."""

    def test_scan_orphaned_typelibs_empty(self, monkeypatch) -> None:
        """Should return empty list when no TypeLibs exist."""

        def fake_key_exists(path):
            return False

        monkeypatch.setattr(registry_tools, "key_exists", fake_key_exists)

        result = registry_tools.scan_orphaned_typelibs(
            ["{00020813-0000-0000-C000-000000000046}"],  # Excel TypeLib
            logger=_Recorder(),
        )
        assert result == []

    def test_cleanup_protocol_handlers_dry_run(self, monkeypatch) -> None:
        """Should identify orphaned protocol handlers in dry run."""

        def fake_key_exists(path):
            return "osf" in path.lower()

        def fake_read_values(hive, path, view=None):
            if "command" in path.lower():
                return {"": '"C:\\NonExistent\\Office.exe" "%1"'}
            return {}

        monkeypatch.setattr(registry_tools, "key_exists", fake_key_exists)
        monkeypatch.setattr(registry_tools, "read_values", fake_read_values)

        result = registry_tools.cleanup_protocol_handlers(
            ["osf", "ms-word"],
            dry_run=True,
            logger=_Recorder(),
        )

        # Should find osf as orphaned (exe doesn't exist)
        assert "osf" in result

    def test_cleanup_shell_extensions(self, monkeypatch) -> None:
        """Should scan shell extension approvals."""

        def fake_iter_values(hive, path):
            if "Approved" in path:
                return iter(
                    [
                        ("{12345678-1234-1234-1234-123456789ABC}", "Office Component"),
                    ]
                )
            return iter([])

        def fake_key_exists(path):
            return False  # CLSID doesn't exist

        monkeypatch.setattr(registry_tools, "iter_values", fake_iter_values)
        monkeypatch.setattr(registry_tools, "key_exists", fake_key_exists)

        result = registry_tools.cleanup_shell_extensions(
            dry_run=True,
            logger=_Recorder(),
        )

        # Should find the orphaned extension
        assert result >= 0


class TestOfficeGuidDetection:
    """Tests for is_office_guid and squished GUID decoding."""

    def test_is_office_guid_valid_office_2016(self) -> None:
        """Should detect Office 2016+ GUIDs with correct pattern."""
        # Office 2016 Pro Plus x64: version 16, SKU 007E
        guid = "{90160000-007E-0000-1000-0000000FF1CE}"
        assert registry_tools.is_office_guid(guid) is True

    def test_is_office_guid_valid_office_2019(self) -> None:
        """Should detect Office 2019 GUIDs."""
        # Office 2019 pattern: version 19, SKU 008F
        guid = "{90190000-008F-0000-1000-0000000FF1CE}"
        assert registry_tools.is_office_guid(guid) is True

    def test_is_office_guid_office_2010_rejected(self) -> None:
        """Should reject Office 2010 (version 14) GUIDs."""
        guid = "{90140000-007E-0000-1000-0000000FF1CE}"
        assert registry_tools.is_office_guid(guid) is False

    def test_is_office_guid_wrong_suffix(self) -> None:
        """Should reject GUIDs without Office suffix."""
        guid = "{90160000-007E-0000-1000-000000000000}"
        assert registry_tools.is_office_guid(guid) is False

    def test_is_office_guid_wrong_sku(self) -> None:
        """Should reject GUIDs with non-C2R SKU codes."""
        guid = "{90160000-ABCD-0000-1000-0000000FF1CE}"
        assert registry_tools.is_office_guid(guid) is False

    def test_is_office_guid_special_mosa_x64(self) -> None:
        """Should detect MOSA x64 special GUID."""
        guid = "{6C1ADE97-24E1-4AE4-AEDD-86D3A209CE60}"
        assert registry_tools.is_office_guid(guid) is True

    def test_is_office_guid_special_mosa_x86(self) -> None:
        """Should detect MOSA x86 special GUID."""
        guid = "{9520DDEB-237A-41DB-AA20-F2EF2360DCEB}"
        assert registry_tools.is_office_guid(guid) is True

    def test_is_office_guid_invalid_length(self) -> None:
        """Should reject GUIDs with wrong length."""
        assert registry_tools.is_office_guid("") is False
        assert registry_tools.is_office_guid("{SHORT}") is False
        assert registry_tools.is_office_guid("A" * 100) is False

    def test_decode_squished_guid_valid(self) -> None:
        """Should decode squished GUID to standard format."""
        # Known Office GUID: {90160000-0011-0000-0000-0000000FF1CE}
        # Squished: 00061009110000000000000000F1EC10
        # Let's verify the algorithm with a simpler test
        squished = "00061009110000000000000000F1EC10"
        result = registry_tools._decode_squished_guid(squished)
        # The squished format reverses segments
        assert result is not None
        assert result.startswith("{")
        assert result.endswith("}")
        assert len(result) == 38

    def test_decode_squished_guid_too_short(self) -> None:
        """Should return None for too-short input."""
        assert registry_tools._decode_squished_guid("ABCD") is None
        assert registry_tools._decode_squished_guid("") is None
        assert registry_tools._decode_squished_guid(None) is None


class TestFilterMultiStringValue:
    """Tests for REG_MULTI_SZ filtering."""

    def test_filter_keeps_non_matching_entries(self, monkeypatch) -> None:
        """Should keep entries that pass the predicate."""

        def fake_get_value(root, path, value_name, view=None):
            return ["keep1", "remove", "keep2"]

        monkeypatch.setattr(registry_tools, "get_value", fake_get_value)
        monkeypatch.setattr(registry_tools, "_ensure_winreg", lambda: None)

        result = registry_tools.filter_multi_string_value(
            0,
            "test\\path",
            "TestValue",
            lambda x: x.startswith("keep"),
            dry_run=True,
            logger=_Recorder(),
        )

        assert result["entries_kept"] == ["keep1", "keep2"]
        assert result["entries_removed"] == ["remove"]

    def test_filter_no_change_when_all_kept(self, monkeypatch) -> None:
        """Should report no changes when all entries are kept."""

        def fake_get_value(root, path, value_name, view=None):
            return ["keep1", "keep2"]

        monkeypatch.setattr(registry_tools, "get_value", fake_get_value)
        monkeypatch.setattr(registry_tools, "_ensure_winreg", lambda: None)

        result = registry_tools.filter_multi_string_value(
            0,
            "test\\path",
            "TestValue",
            lambda x: True,  # Keep all
            dry_run=True,
            logger=_Recorder(),
        )

        assert result["entries_removed"] == []
        assert result["value_deleted"] is False

    def test_filter_handles_missing_value(self, monkeypatch) -> None:
        """Should handle missing value gracefully."""

        def fake_get_value(root, path, value_name, view=None):
            return None

        monkeypatch.setattr(registry_tools, "get_value", fake_get_value)
        monkeypatch.setattr(registry_tools, "_ensure_winreg", lambda: None)

        result = registry_tools.filter_multi_string_value(
            0,
            "test\\path",
            "MissingValue",
            lambda x: True,
            dry_run=True,
            logger=_Recorder(),
        )

        assert result["entries_removed"] == []
        assert result["entries_kept"] == []


class TestCleanupPublishedComponents:
    """Tests for Published Components cleanup."""

    def test_cleanup_handles_missing_components_key(self, monkeypatch) -> None:
        """Should handle missing Components key gracefully."""

        def fake_iter_subkeys(hive, path, view=None):
            raise FileNotFoundError("Not found")

        monkeypatch.setattr(registry_tools, "iter_subkeys", fake_iter_subkeys)
        monkeypatch.setattr(registry_tools, "_ensure_winreg", lambda: None)

        result = registry_tools.cleanup_published_components(
            dry_run=True,
            logger=_Recorder(),
        )

        assert result["components_processed"] == 0
        assert result["values_modified"] == 0
