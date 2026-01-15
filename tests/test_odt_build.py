"""!
@brief Tests for Office Deployment Tool XML configuration builder.
@details Validates ODT XML generation, preset handling, and configuration
options for Office installation scenarios.
"""

from __future__ import annotations

import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest

from office_janitor import odt_build


class TestUpdateChannelEnum:
    """Tests for UpdateChannel enumeration."""

    def test_current_channel_value(self) -> None:
        """Verify Current channel has expected value."""
        assert odt_build.UpdateChannel.CURRENT.value == "Current"

    def test_perpetual_channels_exist(self) -> None:
        """Verify perpetual channels are defined."""
        assert odt_build.UpdateChannel.PERPETUAL_VL_2024.value == "PerpetualVL2024"
        assert odt_build.UpdateChannel.PERPETUAL_VL_2021.value == "PerpetualVL2021"
        assert odt_build.UpdateChannel.PERPETUAL_VL_2019.value == "PerpetualVL2019"

    def test_monthly_enterprise_channel(self) -> None:
        """Verify monthly enterprise channel."""
        assert odt_build.UpdateChannel.MONTHLY_ENTERPRISE.value == "MonthlyEnterprise"


class TestArchitectureEnum:
    """Tests for Architecture enumeration."""

    def test_x64_value(self) -> None:
        """Verify x64 architecture value."""
        assert odt_build.Architecture.X64.value == "64"

    def test_x86_value(self) -> None:
        """Verify x86 architecture value."""
        assert odt_build.Architecture.X86.value == "32"


class TestProductIDs:
    """Tests for PRODUCT_IDS constant."""

    def test_o365_proplus_exists(self) -> None:
        """Verify O365ProPlusRetail product is defined."""
        assert "O365ProPlusRetail" in odt_build.PRODUCT_IDS

    def test_product_has_required_fields(self) -> None:
        """Verify products have required metadata fields."""
        product = odt_build.PRODUCT_IDS["O365ProPlusRetail"]
        assert "name" in product
        assert "channels" in product
        assert "description" in product

    def test_perpetual_products_exist(self) -> None:
        """Verify perpetual Office products are defined."""
        assert "ProPlus2024Volume" in odt_build.PRODUCT_IDS
        assert "ProPlus2021Volume" in odt_build.PRODUCT_IDS
        assert "ProPlus2019Volume" in odt_build.PRODUCT_IDS

    def test_visio_products_exist(self) -> None:
        """Verify Visio products are defined."""
        assert "VisioProRetail" in odt_build.PRODUCT_IDS
        assert "VisioPro2024Volume" in odt_build.PRODUCT_IDS

    def test_project_products_exist(self) -> None:
        """Verify Project products are defined."""
        assert "ProjectProRetail" in odt_build.PRODUCT_IDS
        assert "ProjectPro2024Volume" in odt_build.PRODUCT_IDS


class TestInstallPresets:
    """Tests for INSTALL_PRESETS constant."""

    def test_365_proplus_preset_exists(self) -> None:
        """Verify Microsoft 365 ProPlus preset is defined."""
        assert "365-proplus-x64" in odt_build.INSTALL_PRESETS

    def test_office2024_preset_exists(self) -> None:
        """Verify Office 2024 preset is defined."""
        assert "office2024-x64" in odt_build.INSTALL_PRESETS

    def test_preset_has_required_fields(self) -> None:
        """Verify presets have required configuration fields."""
        preset = odt_build.INSTALL_PRESETS["365-proplus-x64"]
        assert "products" in preset
        assert "architecture" in preset
        assert "channel" in preset
        assert "description" in preset


class TestProductConfig:
    """Tests for ProductConfig dataclass."""

    def test_default_language(self) -> None:
        """Verify default language is en-us."""
        config = odt_build.ProductConfig("O365ProPlusRetail")
        assert config.languages == ["en-us"]

    def test_multiple_languages(self) -> None:
        """Verify multiple languages can be specified."""
        config = odt_build.ProductConfig("O365ProPlusRetail", languages=["en-us", "de-de", "fr-fr"])
        assert len(config.languages) == 3

    def test_exclude_apps(self) -> None:
        """Verify apps can be excluded."""
        config = odt_build.ProductConfig("O365ProPlusRetail", exclude_apps=["OneDrive", "Teams"])
        assert len(config.exclude_apps) == 2

    def test_validate_valid_product(self) -> None:
        """Verify validation passes for valid product."""
        config = odt_build.ProductConfig("O365ProPlusRetail")
        errors = config.validate()
        assert len(errors) == 0

    def test_validate_invalid_product(self) -> None:
        """Verify validation fails for invalid product."""
        config = odt_build.ProductConfig("InvalidProduct")
        errors = config.validate()
        assert len(errors) > 0
        assert any("Unknown product ID" in e for e in errors)


class TestODTConfig:
    """Tests for ODTConfig dataclass."""

    def test_default_values(self) -> None:
        """Verify default configuration values."""
        config = odt_build.ODTConfig()
        assert config.architecture == odt_build.Architecture.X64
        assert config.channel == odt_build.UpdateChannel.CURRENT
        assert config.accept_eula is True
        assert config.force_app_shutdown is True
        assert config.enable_updates is True

    def test_validate_empty_products(self) -> None:
        """Verify validation fails with no products."""
        config = odt_build.ODTConfig()
        errors = config.validate()
        assert len(errors) > 0
        assert any("At least one product" in e for e in errors)

    def test_validate_valid_config(self) -> None:
        """Verify validation passes for valid config."""
        config = odt_build.ODTConfig(products=[odt_build.ProductConfig("O365ProPlusRetail")])
        errors = config.validate()
        assert len(errors) == 0

    def test_from_preset_valid(self) -> None:
        """Verify preset creates valid configuration."""
        config = odt_build.ODTConfig.from_preset("365-proplus-x64")
        assert len(config.products) > 0
        assert config.architecture == odt_build.Architecture.X64
        assert config.channel == odt_build.UpdateChannel.CURRENT

    def test_from_preset_invalid(self) -> None:
        """Verify preset raises error for unknown preset."""
        with pytest.raises(ValueError, match="Unknown preset"):
            odt_build.ODTConfig.from_preset("invalid-preset")

    def test_from_preset_with_languages(self) -> None:
        """Verify preset respects custom languages."""
        config = odt_build.ODTConfig.from_preset("365-proplus-x64", languages=["de-de", "fr-fr"])
        for product in config.products:
            assert product.languages == ["de-de", "fr-fr"]


class TestBuildXml:
    """Tests for build_xml function."""

    def test_basic_xml_structure(self) -> None:
        """Verify basic XML structure is correct."""
        config = odt_build.ODTConfig(products=[odt_build.ProductConfig("O365ProPlusRetail")])
        xml_str = odt_build.build_xml(config)

        # Parse and verify structure
        root = ET.fromstring(xml_str.split("\n", 1)[1])  # Skip XML declaration
        assert root.tag == "Configuration"

        add_elem = root.find("Add")
        assert add_elem is not None
        assert add_elem.get("OfficeClientEdition") == "64"
        assert add_elem.get("Channel") == "Current"

    def test_product_in_xml(self) -> None:
        """Verify product is included in XML."""
        config = odt_build.ODTConfig(products=[odt_build.ProductConfig("O365ProPlusRetail")])
        xml_str = odt_build.build_xml(config)

        root = ET.fromstring(xml_str.split("\n", 1)[1])
        product = root.find(".//Product")
        assert product is not None
        assert product.get("ID") == "O365ProPlusRetail"

    def test_language_in_xml(self) -> None:
        """Verify language is included in XML."""
        config = odt_build.ODTConfig(
            products=[odt_build.ProductConfig("O365ProPlusRetail", languages=["de-de"])]
        )
        xml_str = odt_build.build_xml(config)

        root = ET.fromstring(xml_str.split("\n", 1)[1])
        language = root.find(".//Language")
        assert language is not None
        assert language.get("ID") == "de-de"

    def test_multiple_languages_in_xml(self) -> None:
        """Verify multiple languages are included."""
        config = odt_build.ODTConfig(
            products=[
                odt_build.ProductConfig("O365ProPlusRetail", languages=["en-us", "de-de", "fr-fr"])
            ]
        )
        xml_str = odt_build.build_xml(config)

        root = ET.fromstring(xml_str.split("\n", 1)[1])
        languages = root.findall(".//Language")
        assert len(languages) == 3

    def test_excluded_apps_in_xml(self) -> None:
        """Verify excluded apps are in XML."""
        config = odt_build.ODTConfig(
            products=[
                odt_build.ProductConfig("O365ProPlusRetail", exclude_apps=["OneDrive", "Teams"])
            ]
        )
        xml_str = odt_build.build_xml(config)

        root = ET.fromstring(xml_str.split("\n", 1)[1])
        excludes = root.findall(".//ExcludeApp")
        assert len(excludes) == 2

    def test_updates_element(self) -> None:
        """Verify Updates element is present."""
        config = odt_build.ODTConfig(
            products=[odt_build.ProductConfig("O365ProPlusRetail")],
            enable_updates=True,
        )
        xml_str = odt_build.build_xml(config)

        root = ET.fromstring(xml_str.split("\n", 1)[1])
        updates = root.find("Updates")
        assert updates is not None
        assert updates.get("Enabled") == "TRUE"

    def test_display_element(self) -> None:
        """Verify Display element is present."""
        config = odt_build.ODTConfig(products=[odt_build.ProductConfig("O365ProPlusRetail")])
        xml_str = odt_build.build_xml(config)

        root = ET.fromstring(xml_str.split("\n", 1)[1])
        display = root.find("Display")
        assert display is not None
        assert display.get("Level") == "None"
        assert display.get("AcceptEULA") == "TRUE"

    def test_shared_computer_licensing(self) -> None:
        """Verify shared computer licensing property."""
        config = odt_build.ODTConfig(
            products=[odt_build.ProductConfig("O365ProPlusRetail")],
            shared_computer_licensing=True,
        )
        xml_str = odt_build.build_xml(config)

        root = ET.fromstring(xml_str.split("\n", 1)[1])
        props = root.findall("Property")
        scl_prop = [p for p in props if p.get("Name") == "SharedComputerLicensing"]
        assert len(scl_prop) == 1
        assert scl_prop[0].get("Value") == "1"

    def test_remove_msi_element(self) -> None:
        """Verify RemoveMSI element when enabled."""
        config = odt_build.ODTConfig(
            products=[odt_build.ProductConfig("O365ProPlusRetail")],
            remove_msi=True,
        )
        xml_str = odt_build.build_xml(config)

        root = ET.fromstring(xml_str.split("\n", 1)[1])
        remove_msi = root.find("RemoveMSI")
        assert remove_msi is not None

    def test_invalid_config_raises(self) -> None:
        """Verify invalid config raises ValueError."""
        config = odt_build.ODTConfig()  # No products
        with pytest.raises(ValueError, match="Invalid configuration"):
            odt_build.build_xml(config)

    def test_x86_architecture(self) -> None:
        """Verify x86 architecture in XML."""
        config = odt_build.ODTConfig(
            products=[odt_build.ProductConfig("O365ProPlusRetail")],
            architecture=odt_build.Architecture.X86,
        )
        xml_str = odt_build.build_xml(config)

        root = ET.fromstring(xml_str.split("\n", 1)[1])
        add_elem = root.find("Add")
        assert add_elem.get("OfficeClientEdition") == "32"

    def test_perpetual_channel(self) -> None:
        """Verify perpetual channel in XML."""
        config = odt_build.ODTConfig(
            products=[odt_build.ProductConfig("ProPlus2024Volume")],
            channel=odt_build.UpdateChannel.PERPETUAL_VL_2024,
        )
        xml_str = odt_build.build_xml(config)

        root = ET.fromstring(xml_str.split("\n", 1)[1])
        add_elem = root.find("Add")
        assert add_elem.get("Channel") == "PerpetualVL2024"


class TestBuildRemovalXml:
    """Tests for build_removal_xml function."""

    def test_remove_all_structure(self) -> None:
        """Verify removal XML with Remove All=TRUE."""
        xml_str = odt_build.build_removal_xml(remove_all=True)

        root = ET.fromstring(xml_str.split("\n", 1)[1])
        remove = root.find("Remove")
        assert remove is not None
        assert remove.get("All") == "TRUE"

    def test_remove_specific_products(self) -> None:
        """Verify removal XML with specific products."""
        xml_str = odt_build.build_removal_xml(
            remove_all=False, product_ids=["O365ProPlusRetail", "VisioProRetail"]
        )

        root = ET.fromstring(xml_str.split("\n", 1)[1])
        products = root.findall(".//Product")
        assert len(products) == 2

    def test_remove_msi_element(self) -> None:
        """Verify RemoveMSI in removal XML."""
        xml_str = odt_build.build_removal_xml(remove_msi=True)

        root = ET.fromstring(xml_str.split("\n", 1)[1])
        remove_msi = root.find("RemoveMSI")
        assert remove_msi is not None

    def test_force_app_shutdown(self) -> None:
        """Verify FORCEAPPSHUTDOWN property."""
        xml_str = odt_build.build_removal_xml(force_app_shutdown=True)

        root = ET.fromstring(xml_str.split("\n", 1)[1])
        props = root.findall("Property")
        shutdown_prop = [p for p in props if p.get("Name") == "FORCEAPPSHUTDOWN"]
        assert len(shutdown_prop) == 1


class TestBuildDownloadXml:
    """Tests for build_download_xml function."""

    def test_download_xml_has_source_path(self) -> None:
        """Verify download XML includes SourcePath."""
        config = odt_build.ODTConfig(products=[odt_build.ProductConfig("O365ProPlusRetail")])
        xml_str = odt_build.build_download_xml(config, "C:\\ODTDownload")

        root = ET.fromstring(xml_str.split("\n", 1)[1])
        add_elem = root.find("Add")
        assert add_elem is not None
        assert add_elem.get("SourcePath") == "C:\\ODTDownload"


class TestWriteXmlConfig:
    """Tests for write_xml_config function."""

    def test_write_to_file(self) -> None:
        """Verify XML is written to file."""
        config = odt_build.ODTConfig(products=[odt_build.ProductConfig("O365ProPlusRetail")])

        with tempfile.NamedTemporaryFile(mode="w", suffix=".xml", delete=False) as f:
            output_path = Path(f.name)

        try:
            result = odt_build.write_xml_config(config, output_path)
            assert result == output_path
            assert output_path.exists()

            content = output_path.read_text(encoding="utf-8")
            assert "Configuration" in content
            assert "O365ProPlusRetail" in content
        finally:
            output_path.unlink(missing_ok=True)


class TestWriteTempConfig:
    """Tests for write_temp_config function."""

    def test_creates_temp_file(self) -> None:
        """Verify temp file is created."""
        config = odt_build.ODTConfig(products=[odt_build.ProductConfig("O365ProPlusRetail")])

        temp_path = odt_build.write_temp_config(config)
        try:
            assert temp_path.exists()
            assert temp_path.suffix == ".xml"

            content = temp_path.read_text(encoding="utf-8")
            assert "Configuration" in content
        finally:
            temp_path.unlink(missing_ok=True)


class TestListFunctions:
    """Tests for listing functions."""

    def test_list_products_returns_list(self) -> None:
        """Verify list_products returns a list."""
        products = odt_build.list_products()
        assert isinstance(products, list)
        assert len(products) > 0

    def test_list_products_has_required_fields(self) -> None:
        """Verify products have required fields."""
        products = odt_build.list_products()
        for product in products:
            assert "id" in product
            assert "name" in product
            assert "channels" in product

    def test_list_presets_returns_list(self) -> None:
        """Verify list_presets returns a list."""
        presets = odt_build.list_presets()
        assert isinstance(presets, list)
        assert len(presets) > 0

    def test_list_channels_returns_list(self) -> None:
        """Verify list_channels returns a list."""
        channels = odt_build.list_channels()
        assert isinstance(channels, list)
        assert len(channels) > 0

    def test_list_languages_returns_list(self) -> None:
        """Verify list_languages returns a list."""
        languages = odt_build.list_languages()
        assert isinstance(languages, list)
        assert len(languages) > 0
        assert "en-us" in languages


class TestQuickBuilders:
    """Tests for quick builder functions."""

    def test_build_365_proplus_default(self) -> None:
        """Verify default M365 ProPlus config."""
        config = odt_build.build_365_proplus()
        assert len(config.products) == 1
        assert config.products[0].product_id == "O365ProPlusRetail"
        assert config.architecture == odt_build.Architecture.X64

    def test_build_365_proplus_with_visio(self) -> None:
        """Verify M365 ProPlus with Visio."""
        config = odt_build.build_365_proplus(include_visio=True)
        assert len(config.products) == 2
        product_ids = [p.product_id for p in config.products]
        assert "VisioProRetail" in product_ids

    def test_build_365_proplus_with_project(self) -> None:
        """Verify M365 ProPlus with Project."""
        config = odt_build.build_365_proplus(include_project=True)
        assert len(config.products) == 2
        product_ids = [p.product_id for p in config.products]
        assert "ProjectProRetail" in product_ids

    def test_build_365_proplus_shared_computer(self) -> None:
        """Verify M365 ProPlus with shared computer licensing."""
        config = odt_build.build_365_proplus(shared_computer=True)
        assert config.shared_computer_licensing is True

    def test_build_office_ltsc_2024(self) -> None:
        """Verify Office LTSC 2024 config."""
        config = odt_build.build_office_ltsc("2024")
        assert len(config.products) == 1
        assert config.products[0].product_id == "ProPlus2024Volume"
        assert config.channel == odt_build.UpdateChannel.PERPETUAL_VL_2024

    def test_build_office_ltsc_2021(self) -> None:
        """Verify Office LTSC 2021 config."""
        config = odt_build.build_office_ltsc("2021")
        assert config.products[0].product_id == "ProPlus2021Volume"
        assert config.channel == odt_build.UpdateChannel.PERPETUAL_VL_2021

    def test_build_office_ltsc_2019(self) -> None:
        """Verify Office LTSC 2019 config."""
        config = odt_build.build_office_ltsc("2019")
        assert config.products[0].product_id == "ProPlus2019Volume"
        assert config.channel == odt_build.UpdateChannel.PERPETUAL_VL_2019

    def test_build_office_ltsc_retail(self) -> None:
        """Verify Office LTSC retail config."""
        config = odt_build.build_office_ltsc("2024", volume=False)
        assert config.products[0].product_id == "ProPlus2024Retail"

    def test_build_office_ltsc_invalid_version(self) -> None:
        """Verify error for invalid LTSC version."""
        with pytest.raises(ValueError, match="Unsupported LTSC version"):
            odt_build.build_office_ltsc("2016")


class TestSupportedLanguages:
    """Tests for SUPPORTED_LANGUAGES constant."""

    def test_common_languages_present(self) -> None:
        """Verify common languages are in the list."""
        assert "en-us" in odt_build.SUPPORTED_LANGUAGES
        assert "de-de" in odt_build.SUPPORTED_LANGUAGES
        assert "fr-fr" in odt_build.SUPPORTED_LANGUAGES
        assert "es-es" in odt_build.SUPPORTED_LANGUAGES
        assert "ja-jp" in odt_build.SUPPORTED_LANGUAGES
        assert "zh-cn" in odt_build.SUPPORTED_LANGUAGES

    def test_languages_are_lowercase(self) -> None:
        """Verify all language codes are lowercase."""
        for lang in odt_build.SUPPORTED_LANGUAGES:
            assert lang == lang.lower()


class TestLTSCFullPresets:
    """Tests for the new LTSC full presets with Visio + Project."""

    def test_ltsc2024_full_preset_exists(self) -> None:
        """Verify LTSC 2024 full preset is defined."""
        assert "ltsc2024-full-x64" in odt_build.INSTALL_PRESETS
        assert "ltsc2024-full-x86" in odt_build.INSTALL_PRESETS

    def test_ltsc2021_full_preset_exists(self) -> None:
        """Verify LTSC 2021 full preset is defined."""
        assert "ltsc2021-full-x64" in odt_build.INSTALL_PRESETS

    def test_ltsc2024_full_has_all_products(self) -> None:
        """Verify LTSC 2024 full preset includes ProPlus, Visio, and Project."""
        preset = odt_build.INSTALL_PRESETS["ltsc2024-full-x64"]
        products = preset["products"]
        assert "ProPlus2024Volume" in products
        assert "VisioPro2024Volume" in products
        assert "ProjectPro2024Volume" in products

    def test_ltsc2024_full_preset_config(self) -> None:
        """Verify LTSC 2024 full preset generates correct config."""
        config = odt_build.ODTConfig.from_preset("ltsc2024-full-x64", ["en-us", "es-mx"])
        assert len(config.products) == 3
        product_ids = [p.product_id for p in config.products]
        assert "ProPlus2024Volume" in product_ids
        assert "VisioPro2024Volume" in product_ids
        assert "ProjectPro2024Volume" in product_ids
        assert config.channel == odt_build.UpdateChannel.PERPETUAL_VL_2024
        assert config.architecture == odt_build.Architecture.X64
        # Check languages are applied to all products
        for p in config.products:
            assert "en-us" in p.languages
            assert "es-mx" in p.languages


class TestGetOdtSetupPath:
    """Tests for ODT setup.exe path discovery."""

    def test_get_odt_setup_path_finds_exe(self) -> None:
        """Verify setup.exe can be found in oem folder."""
        # This test assumes setup.exe exists in oem/
        try:
            path = odt_build.get_odt_setup_path()
            assert path.name == "setup.exe"
            assert path.exists()
        except FileNotFoundError:
            pytest.skip("setup.exe not present in oem folder")


class TestODTResult:
    """Tests for ODTResult dataclass."""

    def test_odt_result_defaults(self) -> None:
        """Verify ODTResult has correct default values."""
        result = odt_build.ODTResult(
            success=True,
            return_code=0,
            command=["test"],
            config_path=None,
            stdout="",
            stderr="",
            duration=1.5,
        )
        assert result.success is True
        assert result.error is None

    def test_odt_result_with_error(self) -> None:
        """Verify ODTResult captures error information."""
        result = odt_build.ODTResult(
            success=False,
            return_code=1,
            command=["setup.exe", "/configure", "config.xml"],
            config_path=Path("config.xml"),
            stdout="",
            stderr="Error occurred",
            duration=2.0,
            error="Installation failed",
        )
        assert result.success is False
        assert result.return_code == 1
        assert result.error == "Installation failed"


class TestInstallLtsc2024Full:
    """Tests for install_ltsc_2024_full quick function."""

    def test_install_ltsc_2024_full_config(self) -> None:
        """Verify install_ltsc_2024_full generates correct config internally."""
        # We can't actually run the install, but we can test the config generation
        config = odt_build.build_office_ltsc(
            "2024",
            languages=["en-us", "es-mx", "pt-br"],
            volume=True,
            include_visio=True,
            include_project=True,
        )
        assert len(config.products) == 3
        product_ids = [p.product_id for p in config.products]
        assert "ProPlus2024Volume" in product_ids
        assert "VisioPro2024Volume" in product_ids
        assert "ProjectPro2024Volume" in product_ids
        # Check languages
        for p in config.products:
            assert "en-us" in p.languages
            assert "es-mx" in p.languages
            assert "pt-br" in p.languages


class TestProgressMonitoring:
    """Tests for ODT progress monitoring functions."""

    def test_get_odt_log_path(self) -> None:
        """Verify log path returns valid directory."""
        log_path = odt_build._get_odt_log_path()
        assert isinstance(log_path, Path)
        # Should be TEMP directory or similar
        assert log_path.exists() or log_path.parent.exists()

    def test_find_latest_odt_log_no_logs(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """Verify returns None when no logs exist."""
        monkeypatch.setattr(odt_build, "_get_odt_log_path", lambda: tmp_path)
        result = odt_build._find_latest_odt_log()
        assert result is None

    def test_find_latest_odt_log_with_logs(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """Verify returns latest log file when logs exist."""
        monkeypatch.setattr(odt_build, "_get_odt_log_path", lambda: tmp_path)
        # Create mock log files
        log1 = tmp_path / "Microsoft Office Click-to-Run 1.log"
        log2 = tmp_path / "Microsoft Office Click-to-Run 2.log"
        log1.write_text("old log")
        import time
        time.sleep(0.01)  # Ensure different timestamps
        log2.write_text("new log")
        
        result = odt_build._find_latest_odt_log()
        assert result == log2

    def test_parse_odt_progress_nonexistent_file(self) -> None:
        """Verify returns default message for nonexistent file."""
        status, pct = odt_build._parse_odt_progress(Path("/nonexistent/log.txt"))
        assert status == "Starting installation..."
        assert pct is None

    def test_parse_odt_progress_downloading(self, tmp_path: Path) -> None:
        """Verify parses downloading status."""
        log_file = tmp_path / "test.log"
        log_file.write_text("2024-01-01 downloading files 50%\n")
        
        status, pct = odt_build._parse_odt_progress(log_file)
        assert "Downloading" in status
        assert pct == 50

    def test_parse_odt_progress_installing(self, tmp_path: Path) -> None:
        """Verify parses installing status."""
        log_file = tmp_path / "test.log"
        log_file.write_text("2024-01-01 Installing Office components 75%\n")
        
        status, pct = odt_build._parse_odt_progress(log_file)
        assert "Installing" in status
        assert pct == 75

    def test_parse_odt_progress_configuring(self, tmp_path: Path) -> None:
        """Verify parses configuring status."""
        log_file = tmp_path / "test.log"
        log_file.write_text("2024-01-01 Configuring Office settings\n")
        
        status, pct = odt_build._parse_odt_progress(log_file)
        assert "Configuring" in status
        assert pct is None

    def test_parse_odt_progress_percentage_only(self, tmp_path: Path) -> None:
        """Verify parses percentage without specific context."""
        log_file = tmp_path / "test.log"
        log_file.write_text("Progress: 80%\n")
        
        status, pct = odt_build._parse_odt_progress(log_file)
        assert pct == 80


class TestInstallMetrics:
    """Tests for installation metrics functions."""

    def test_format_size_bytes(self) -> None:
        """Verify format_size handles bytes."""
        assert odt_build._format_size(500) == "500 B"

    def test_format_size_kilobytes(self) -> None:
        """Verify format_size handles kilobytes."""
        result = odt_build._format_size(2048)
        assert "KB" in result

    def test_format_size_megabytes(self) -> None:
        """Verify format_size handles megabytes."""
        result = odt_build._format_size(50 * 1024 * 1024)
        assert "MB" in result
        assert "50" in result

    def test_format_size_gigabytes(self) -> None:
        """Verify format_size handles gigabytes."""
        result = odt_build._format_size(2 * 1024 * 1024 * 1024)
        assert "GB" in result
        assert "2" in result

    def test_get_folder_size_empty(self, tmp_path: Path) -> None:
        """Verify get_folder_size returns 0 for empty folder."""
        size = odt_build._get_folder_size(tmp_path)
        assert size == 0

    def test_get_folder_size_with_files(self, tmp_path: Path) -> None:
        """Verify get_folder_size counts file sizes."""
        # Create test files
        (tmp_path / "file1.txt").write_text("hello")
        (tmp_path / "file2.txt").write_text("world!")
        
        size = odt_build._get_folder_size(tmp_path)
        assert size > 0
        assert size >= 11  # At least "hello" + "world!"

    def test_get_folder_size_nonexistent(self) -> None:
        """Verify get_folder_size returns 0 for nonexistent folder."""
        size = odt_build._get_folder_size(Path("/nonexistent/path"))
        assert size == 0

    def test_count_office_files_returns_int(self) -> None:
        """Verify count_office_files returns an integer."""
        count = odt_build._count_office_files()
        assert isinstance(count, int)
        assert count >= 0

    def test_capture_install_metrics_returns_dataclass(self) -> None:
        """Verify capture_install_metrics returns proper dataclass."""
        metrics = odt_build._capture_install_metrics()
        assert isinstance(metrics, odt_build.InstallMetrics)
        assert isinstance(metrics.install_size, int)
        assert isinstance(metrics.file_count, int)
        assert isinstance(metrics.registry_keys, int)
        assert metrics.log_status  # Should have some status string
