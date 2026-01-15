"""!
@brief Tests for the Office repair module.
@details Validates repair configuration, detection, and command generation
without executing actual repair operations.
"""

from __future__ import annotations

import re
from pathlib import Path
from unittest import mock

import pytest

from office_janitor import repair
from office_janitor.repair import (
    DEFAULT_CULTURE,
    PLATFORM_X64,
    PLATFORM_X86,
    REPAIR_TIMEOUT_FULL,
    REPAIR_TIMEOUT_QUICK,
    SUPPORTED_CULTURES,
    DisplayLevel,
    RepairConfig,
    RepairResult,
    RepairType,
)

# ---------------------------------------------------------------------------
# RepairConfig Tests
# ---------------------------------------------------------------------------


class TestRepairConfig:
    """Tests for RepairConfig dataclass and factory methods."""

    def test_default_values(self) -> None:
        """Test default configuration values."""
        config = RepairConfig()
        assert config.repair_type == RepairType.QUICK
        assert config.platform == PLATFORM_X64
        assert config.culture == DEFAULT_CULTURE
        assert config.force_app_shutdown is True
        assert config.display_level == DisplayLevel.SILENT
        assert config.timeout is None

    def test_quick_repair_factory(self) -> None:
        """Test quick_repair factory method."""
        config = RepairConfig.quick_repair()
        assert config.repair_type == RepairType.QUICK
        assert config.force_app_shutdown is True
        assert config.display_level == DisplayLevel.SILENT

    def test_full_repair_factory(self) -> None:
        """Test full_repair factory method."""
        config = RepairConfig.full_repair()
        assert config.repair_type == RepairType.FULL
        assert config.force_app_shutdown is True
        assert config.display_level == DisplayLevel.SILENT

    def test_quick_repair_with_custom_values(self) -> None:
        """Test quick_repair with custom parameters."""
        config = RepairConfig.quick_repair(
            platform=PLATFORM_X86,
            culture="de-de",
            force_shutdown=False,
            silent=False,
        )
        assert config.platform == PLATFORM_X86
        assert config.culture == "de-de"
        assert config.force_app_shutdown is False
        assert config.display_level == DisplayLevel.VISIBLE

    def test_full_repair_with_custom_values(self) -> None:
        """Test full_repair with custom parameters."""
        config = RepairConfig.full_repair(
            platform=PLATFORM_X64,
            culture="fr-fr",
            force_shutdown=True,
            silent=True,
        )
        assert config.platform == PLATFORM_X64
        assert config.culture == "fr-fr"
        assert config.force_app_shutdown is True
        assert config.display_level == DisplayLevel.SILENT

    def test_effective_timeout_quick(self) -> None:
        """Test effective timeout for quick repair."""
        config = RepairConfig.quick_repair()
        assert config.effective_timeout == REPAIR_TIMEOUT_QUICK

    def test_effective_timeout_full(self) -> None:
        """Test effective timeout for full repair."""
        config = RepairConfig.full_repair()
        assert config.effective_timeout == REPAIR_TIMEOUT_FULL

    def test_explicit_timeout_overrides_default(self) -> None:
        """Test that explicit timeout takes precedence."""
        config = RepairConfig(timeout=300)
        assert config.effective_timeout == 300

    def test_to_command_args(self) -> None:
        """Test command argument generation."""
        config = RepairConfig(
            repair_type=RepairType.QUICK,
            platform=PLATFORM_X64,
            culture="en-us",
            force_app_shutdown=True,
            display_level=DisplayLevel.SILENT,
        )
        args = config.to_command_args()
        assert "scenario=Repair" in args
        assert "platform=x64" in args
        assert "culture=en-us" in args
        assert "RepairType=QuickRepair" in args
        assert "forceappshutdown=True" in args
        assert "DisplayLevel=False" in args

    def test_to_command_args_full_repair(self) -> None:
        """Test command arguments for full repair."""
        config = RepairConfig.full_repair(silent=False)
        args = config.to_command_args()
        assert "RepairType=FullRepair" in args
        assert "DisplayLevel=True" in args


class TestRepairConfigValidation:
    """Tests for RepairConfig validation."""

    def test_valid_culture_codes(self) -> None:
        """Test that standard culture codes are accepted."""
        for culture in ["en-us", "de-de", "ja-jp", "zh-cn"]:
            config = RepairConfig(culture=culture)
            assert config.culture == culture

    def test_culture_from_supported_list(self) -> None:
        """Test all cultures in SUPPORTED_CULTURES."""
        for culture in SUPPORTED_CULTURES[:5]:  # Test a sample
            config = RepairConfig(culture=culture)
            assert config.culture == culture

    def test_unknown_culture_with_valid_format(self) -> None:
        """Test that unknown cultures with valid format are accepted."""
        config = RepairConfig(culture="xx-yy")
        assert config.culture == "xx-yy"

    def test_invalid_culture_format_raises(self) -> None:
        """Test that invalid culture format raises ValueError."""
        with pytest.raises(ValueError, match="Invalid culture code"):
            RepairConfig(culture="invalid")


# ---------------------------------------------------------------------------
# RepairResult Tests
# ---------------------------------------------------------------------------


class TestRepairResult:
    """Tests for RepairResult dataclass."""

    def test_success_result(self) -> None:
        """Test successful repair result."""
        result = RepairResult(
            success=True,
            repair_type=RepairType.QUICK,
            return_code=0,
            duration=120.5,
        )
        assert result.success is True
        assert "successfully" in result.summary.lower()
        assert "120.5s" in result.summary

    def test_failed_result(self) -> None:
        """Test failed repair result."""
        result = RepairResult(
            success=False,
            repair_type=RepairType.FULL,
            return_code=1603,
            duration=30.0,
            error_message="Installation failed",
        )
        assert result.success is False
        assert "failed" in result.summary.lower()
        assert "1603" in result.summary

    def test_skipped_result(self) -> None:
        """Test skipped (dry-run) repair result."""
        result = RepairResult(
            success=True,
            repair_type=RepairType.QUICK,
            return_code=0,
            duration=0.0,
            skipped=True,
        )
        assert result.skipped is True
        assert "skipped" in result.summary.lower()
        assert "dry-run" in result.summary.lower()

    def test_timed_out_result(self) -> None:
        """Test timed-out repair result."""
        result = RepairResult(
            success=False,
            repair_type=RepairType.FULL,
            return_code=-1,
            duration=3600.0,
            timed_out=True,
        )
        assert result.timed_out is True
        assert "timed out" in result.summary.lower()


# ---------------------------------------------------------------------------
# Detection Function Tests
# ---------------------------------------------------------------------------


class TestDetectionFunctions:
    """Tests for Office detection utilities."""

    def test_detect_office_platform_default(self) -> None:
        """Test platform detection returns valid value."""
        # Mock registry to return None (no detection)
        with mock.patch(
            "office_janitor.repair.registry_tools.get_value",
            return_value=None,
        ):
            platform = repair._detect_office_platform()
            assert platform in (PLATFORM_X86, PLATFORM_X64)

    def test_detect_office_platform_x86(self) -> None:
        """Test platform detection for x86 Office."""
        with mock.patch(
            "office_janitor.repair.registry_tools.get_value",
            return_value="x86",
        ):
            platform = repair._detect_office_platform()
            assert platform == PLATFORM_X86

    def test_detect_office_platform_x64(self) -> None:
        """Test platform detection for x64 Office."""
        with mock.patch(
            "office_janitor.repair.registry_tools.get_value",
            return_value="x64",
        ):
            platform = repair._detect_office_platform()
            assert platform == PLATFORM_X64

    def test_detect_office_culture_default(self) -> None:
        """Test culture detection falls back to default."""
        with mock.patch(
            "office_janitor.repair.registry_tools.get_value",
            return_value=None,
        ):
            culture = repair._detect_office_culture()
            # Should return en-us or system locale
            assert re.match(r"^[a-z]{2}-[a-z]{2}$", culture)

    def test_detect_office_culture_from_registry(self) -> None:
        """Test culture detection from registry."""
        with mock.patch(
            "office_janitor.repair.registry_tools.get_value",
            return_value="de-DE",
        ):
            culture = repair._detect_office_culture()
            assert culture == "de-de"


class TestIsC2ROfficeInstalled:
    """Tests for C2R Office detection."""

    def test_c2r_detected(self) -> None:
        """Test detection when C2R is installed."""
        with mock.patch(
            "office_janitor.repair.registry_tools.get_value",
            return_value="x64",
        ):
            assert repair.is_c2r_office_installed() is True

    def test_c2r_not_detected(self) -> None:
        """Test detection when C2R is not installed."""
        with mock.patch(
            "office_janitor.repair.registry_tools.get_value",
            return_value=None,
        ):
            assert repair.is_c2r_office_installed() is False

    def test_c2r_detection_exception(self) -> None:
        """Test detection handles exceptions gracefully."""
        with mock.patch(
            "office_janitor.repair.registry_tools.get_value",
            side_effect=Exception("Registry error"),
        ):
            assert repair.is_c2r_office_installed() is False


class TestGetInstalledC2RInfo:
    """Tests for getting installed C2R information."""

    def test_full_info_available(self) -> None:
        """Test retrieving full installation info."""

        def mock_read(hive: int, path: str, name: str) -> str | None:
            values = {
                "VersionToReport": "16.0.14326.20454",
                "Platform": "x64",
                "ClientCulture": "en-us",
                "ProductReleaseIds": "O365ProPlusRetail",
                "CDNBaseUrl": "http://officecdn.microsoft.com/",
            }
            return values.get(name)

        with mock.patch(
            "office_janitor.repair.registry_tools.get_value",
            side_effect=mock_read,
        ):
            info = repair.get_installed_c2r_info()
            assert info["version"] == "16.0.14326.20454"
            assert info["platform"] == "x64"
            assert info["culture"] == "en-us"
            assert info["product_ids"] == "O365ProPlusRetail"

    def test_info_handles_missing_values(self) -> None:
        """Test info retrieval handles missing registry values."""
        with mock.patch(
            "office_janitor.repair.registry_tools.get_value",
            return_value=None,
        ):
            info = repair.get_installed_c2r_info()
            assert info["version"] is None
            assert info["platform"] is None


# ---------------------------------------------------------------------------
# Executable Location Tests
# ---------------------------------------------------------------------------


class TestFindOfficeClickToRunExe:
    """Tests for locating OfficeClickToRun.exe."""

    def test_custom_path_exists(self, tmp_path: Path) -> None:
        """Test custom path takes precedence when it exists."""
        custom_exe = tmp_path / "OfficeClickToRun.exe"
        custom_exe.touch()
        result = repair.find_officeclicktorun_exe(custom_exe)
        assert result == custom_exe

    def test_custom_path_not_exists(self, tmp_path: Path) -> None:
        """Test custom path fallback when it doesn't exist."""
        custom_exe = tmp_path / "nonexistent.exe"
        # When custom path doesn't exist and system paths don't exist,
        # returns None (or bundled if available)
        result = repair.find_officeclicktorun_exe(custom_exe)
        # Result depends on system state
        assert result is None or isinstance(result, Path)


class TestFindOdtSetupExe:
    """Tests for locating ODT setup.exe."""

    def test_custom_path_exists(self, tmp_path: Path) -> None:
        """Test custom path takes precedence when it exists."""
        custom_exe = tmp_path / "setup.exe"
        custom_exe.touch()
        result = repair.find_odt_setup_exe(custom_exe)
        assert result == custom_exe


# ---------------------------------------------------------------------------
# Repair Execution Tests (Mocked)
# ---------------------------------------------------------------------------


class TestRunRepair:
    """Tests for repair execution with mocked subprocess."""

    def test_dry_run_skips_execution(self) -> None:
        """Test that dry-run mode skips actual execution."""
        config = RepairConfig.quick_repair()

        with mock.patch(
            "office_janitor.repair.find_officeclicktorun_exe",
            return_value=Path("C:/fake/OfficeClickToRun.exe"),
        ):
            result = repair.run_repair(config, dry_run=True)
            assert result.skipped is True
            assert result.success is True

    def test_exe_not_found_error(self) -> None:
        """Test error handling when executable not found."""
        config = RepairConfig.quick_repair()

        with mock.patch(
            "office_janitor.repair.find_officeclicktorun_exe",
            return_value=None,
        ):
            result = repair.run_repair(config, dry_run=False)
            assert result.success is False
            assert "not found" in result.error_message.lower()

    def test_successful_repair(self) -> None:
        """Test successful repair execution."""
        config = RepairConfig.quick_repair()

        mock_cmd_result = mock.MagicMock()
        mock_cmd_result.returncode = 0
        mock_cmd_result.stdout = ""
        mock_cmd_result.stderr = ""
        mock_cmd_result.duration = 60.0
        mock_cmd_result.timed_out = False
        mock_cmd_result.error = None

        with mock.patch(
            "office_janitor.repair.find_officeclicktorun_exe",
            return_value=Path("C:/fake/OfficeClickToRun.exe"),
        ):
            with mock.patch(
                "office_janitor.repair.command_runner.run_command",
                return_value=mock_cmd_result,
            ):
                with mock.patch("office_janitor.repair._close_office_applications"):
                    result = repair.run_repair(config, dry_run=False)
                    assert result.success is True
                    assert result.return_code == 0


class TestConvenienceFunctions:
    """Tests for quick_repair and full_repair convenience functions."""

    def test_quick_repair_default(self) -> None:
        """Test quick_repair with defaults."""
        with mock.patch("office_janitor.repair.run_repair") as mock_run:
            mock_run.return_value = RepairResult(
                success=True,
                repair_type=RepairType.QUICK,
                return_code=0,
                duration=60.0,
                skipped=True,
            )
            result = repair.quick_repair(dry_run=True)
            assert result.repair_type == RepairType.QUICK
            mock_run.assert_called_once()

    def test_full_repair_default(self) -> None:
        """Test full_repair with defaults."""
        with mock.patch("office_janitor.repair.run_repair") as mock_run:
            mock_run.return_value = RepairResult(
                success=True,
                repair_type=RepairType.FULL,
                return_code=0,
                duration=120.0,
                skipped=True,
            )
            result = repair.full_repair(dry_run=True)
            assert result.repair_type == RepairType.FULL
            mock_run.assert_called_once()


# ---------------------------------------------------------------------------
# XML Configuration Tests
# ---------------------------------------------------------------------------


class TestGenerateRepairConfigXml:
    """Tests for XML configuration generation."""

    def test_generates_valid_xml(self, tmp_path: Path) -> None:
        """Test XML file generation."""
        output = tmp_path / "repair.xml"

        with mock.patch(
            "office_janitor.repair.get_installed_c2r_info",
            return_value={"product_ids": "O365ProPlusRetail"},
        ):
            result = repair.generate_repair_config_xml(output)

            assert result.exists()
            content = result.read_text()
            assert "<Configuration>" in content
            assert "O365ProPlusRetail" in content
            assert "FORCEAPPSHUTDOWN" in content

    def test_custom_product_ids(self, tmp_path: Path) -> None:
        """Test XML with custom product IDs."""
        output = tmp_path / "custom.xml"

        result = repair.generate_repair_config_xml(
            output,
            product_ids=["VisioProRetail", "ProjectProRetail"],
            language="de-de",
        )

        content = result.read_text()
        assert "VisioProRetail" in content
        assert "ProjectProRetail" in content
        assert "de-de" in content

    def test_creates_parent_directory(self, tmp_path: Path) -> None:
        """Test that parent directories are created."""
        output = tmp_path / "nested" / "dir" / "repair.xml"

        with mock.patch(
            "office_janitor.repair.get_installed_c2r_info",
            return_value={"product_ids": "O365ProPlusRetail"},
        ):
            result = repair.generate_repair_config_xml(output)
            assert result.exists()


class TestReconfigureOffice:
    """Tests for ODT-based reconfiguration."""

    def test_config_file_not_found(self, tmp_path: Path) -> None:
        """Test error when config file doesn't exist."""
        config_path = tmp_path / "nonexistent.xml"

        with mock.patch(
            "office_janitor.repair.find_odt_setup_exe",
            return_value=Path("C:/fake/setup.exe"),
        ):
            result = repair.reconfigure_office(config_path)
            assert result.returncode == -1
            assert "not found" in result.error.lower()

    def test_setup_exe_not_found(self, tmp_path: Path) -> None:
        """Test error when setup.exe not found."""
        config_path = tmp_path / "config.xml"
        config_path.touch()

        with mock.patch(
            "office_janitor.repair.find_odt_setup_exe",
            return_value=None,
        ):
            result = repair.reconfigure_office(config_path)
            assert result.returncode == -1
            assert "not found" in result.stderr.lower()


# ---------------------------------------------------------------------------
# Constants Tests
# ---------------------------------------------------------------------------


class TestConstants:
    """Tests for module constants."""

    def test_supported_cultures_not_empty(self) -> None:
        """Test SUPPORTED_CULTURES is populated."""
        assert len(SUPPORTED_CULTURES) > 0

    def test_supported_cultures_format(self) -> None:
        """Test all cultures match expected format."""
        # Pattern allows for standard ll-cc or extended ll-cccc-cc formats
        pattern = re.compile(r"^[a-z]{2,3}(-[a-z]{2,6})?(-[a-z]{2})?$")
        for culture in SUPPORTED_CULTURES:
            assert pattern.match(culture), f"Invalid culture format: {culture}"

    def test_timeouts_are_positive(self) -> None:
        """Test timeout constants are positive integers."""
        assert REPAIR_TIMEOUT_QUICK > 0
        assert REPAIR_TIMEOUT_FULL > 0
        assert REPAIR_TIMEOUT_FULL > REPAIR_TIMEOUT_QUICK

    def test_platforms_are_valid(self) -> None:
        """Test platform constants."""
        assert PLATFORM_X86 == "x86"
        assert PLATFORM_X64 == "x64"

    def test_default_culture(self) -> None:
        """Test default culture is set."""
        assert DEFAULT_CULTURE == "en-us"


# ---------------------------------------------------------------------------
# RepairType Enum Tests
# ---------------------------------------------------------------------------


class TestRepairTypeEnum:
    """Tests for RepairType enumeration."""

    def test_quick_value(self) -> None:
        """Test QuickRepair enum value."""
        assert RepairType.QUICK.value == "QuickRepair"

    def test_full_value(self) -> None:
        """Test FullRepair enum value."""
        assert RepairType.FULL.value == "FullRepair"


class TestDisplayLevelEnum:
    """Tests for DisplayLevel enumeration."""

    def test_silent_value(self) -> None:
        """Test silent display level."""
        assert DisplayLevel.SILENT.value == "False"

    def test_visible_value(self) -> None:
        """Test visible display level."""
        assert DisplayLevel.VISIBLE.value == "True"


# ---------------------------------------------------------------------------
# OEM Config Tests
# ---------------------------------------------------------------------------


class TestOemConfigPresets:
    """Tests for OEM configuration preset functionality."""

    def test_presets_dictionary_not_empty(self) -> None:
        """Test OEM_CONFIG_PRESETS is populated."""
        from office_janitor.repair import OEM_CONFIG_PRESETS

        assert len(OEM_CONFIG_PRESETS) > 0

    def test_presets_have_xml_files(self) -> None:
        """Test all presets point to XML files."""
        from office_janitor.repair import OEM_CONFIG_PRESETS

        for name, filename in OEM_CONFIG_PRESETS.items():
            assert filename.endswith(".xml"), f"Preset {name} should reference XML file"

    def test_presets_expected_entries(self) -> None:
        """Test expected preset names exist."""
        from office_janitor.repair import OEM_CONFIG_PRESETS

        expected = [
            "full-removal",
            "quick-repair",
            "full-repair",
            "proplus-x64",
        ]
        for name in expected:
            assert name in OEM_CONFIG_PRESETS, f"Expected preset '{name}' not found"


class TestGetOemConfigPath:
    """Tests for get_oem_config_path function."""

    def test_returns_none_for_nonexistent_preset(self) -> None:
        """Test returns None for unknown preset."""
        from office_janitor.repair import get_oem_config_path

        result = get_oem_config_path("nonexistent-preset-xyz")
        assert result is None

    def test_accepts_preset_name(self) -> None:
        """Test resolves preset name from dictionary."""
        from office_janitor.repair import OEM_CONFIG_PRESETS, get_oem_config_path

        # Test with first preset that exists in the mapping
        preset_name = list(OEM_CONFIG_PRESETS.keys())[0]
        result = get_oem_config_path(preset_name)
        # Result may be None if file doesn't exist, but function shouldn't raise
        # Just verify it runs without error
        assert result is None or result.name.endswith(".xml")

    def test_accepts_absolute_path(self, tmp_path: Path) -> None:
        """Test accepts direct absolute path to XML file."""
        from office_janitor.repair import get_oem_config_path

        # Create a temporary XML file
        xml_file = tmp_path / "test_config.xml"
        xml_file.write_text("<Configuration />", encoding="utf-8")

        result = get_oem_config_path(str(xml_file))
        assert result is not None
        assert result == xml_file


class TestListOemConfigs:
    """Tests for list_oem_configs function."""

    def test_returns_list(self) -> None:
        """Test returns a list."""
        from office_janitor.repair import list_oem_configs

        result = list_oem_configs()
        assert isinstance(result, list)

    def test_returns_tuples_with_correct_format(self) -> None:
        """Test returned items are (name, filename, exists) tuples."""
        from office_janitor.repair import list_oem_configs

        result = list_oem_configs()
        for item in result:
            assert len(item) == 3
            name, filename, exists = item
            assert isinstance(name, str)
            assert isinstance(filename, str)
            assert isinstance(exists, bool)

    def test_includes_all_presets(self) -> None:
        """Test all presets from OEM_CONFIG_PRESETS are listed."""
        from office_janitor.repair import OEM_CONFIG_PRESETS, list_oem_configs

        result = list_oem_configs()
        listed_names = {item[0] for item in result}
        for name in OEM_CONFIG_PRESETS:
            assert name in listed_names


class TestRunOemConfig:
    """Tests for run_oem_config function."""

    def test_returns_error_for_nonexistent_preset(self) -> None:
        """Test returns error for unknown preset."""
        from office_janitor.repair import run_oem_config

        result = run_oem_config("nonexistent-preset-xyz", dry_run=True)
        assert result.returncode != 0
        assert "not found" in (result.error or result.stderr or "").lower()

    def test_calls_reconfigure_office(self, tmp_path: Path) -> None:
        """Test passes config path to reconfigure_office."""
        from office_janitor.exec_utils import CommandResult
        from office_janitor.repair import run_oem_config

        config_path = tmp_path / "test.xml"

        with mock.patch("office_janitor.repair.get_oem_config_path") as mock_get_path:
            with mock.patch("office_janitor.repair.reconfigure_office") as mock_reconfigure:
                mock_get_path.return_value = config_path
                mock_reconfigure.return_value = CommandResult(
                    command=["test"],
                    returncode=0,
                    stdout="success",
                    stderr="",
                    duration=1.0,
                )

                run_oem_config("test-preset", dry_run=True)

                mock_reconfigure.assert_called_once()
                call_args = mock_reconfigure.call_args
                assert call_args[0][0] == config_path
                assert call_args[1]["dry_run"] is True
