"""!
@brief Tests for MSI Component Scanner with mocked Windows Installer COM.
@details Validates the component scanning logic without requiring actual
Windows Installer access.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any
from unittest.mock import MagicMock, patch

import pytest

from office_janitor.msi_components import (
    ComponentInfo,
    MSIComponentScanner,
    MsiInstallState,
    ProductInfo,
    ScanResult,
    WindowsInstallerError,
    enumerate_components,
    enumerate_products,
    get_component_clients,
    get_component_path,
    get_component_state,
    is_office_component,
)


class MockInstaller:
    """Mock Windows Installer COM object for testing."""

    def __init__(self) -> None:
        self._products: dict[str, dict[str, str]] = {}
        self._components: dict[str, list[str]] = {}  # component -> [products]
        self._component_paths: dict[tuple[str, str], str] = {}  # (product, comp) -> path

    def add_product(
        self,
        product_code: str,
        name: str = "",
        version: str = "",
        state: int = 5,
    ) -> None:
        """Add a mock product."""
        self._products[product_code] = {
            "ProductName": name,
            "VersionString": version,
            "InstallLocation": r"C:\Program Files\Test",
            "InstallSource": r"C:\Temp",
            "PackageCode": "{00000000-0000-0000-0000-000000000000}",
        }

    def add_component(
        self,
        component_id: str,
        clients: list[str],
        paths: dict[str, str] | None = None,
    ) -> None:
        """Add a mock component."""
        self._components[component_id] = clients
        if paths:
            for product, path in paths.items():
                self._component_paths[(product, component_id)] = path

    @property
    def Products(self) -> list[str]:
        """Return list of product codes."""
        return list(self._products.keys())

    @property
    def Components(self) -> list[str]:
        """Return list of component IDs."""
        return list(self._components.keys())

    def ProductInfo(self, product_code: str, prop: str) -> str:
        """Get product property."""
        if product_code in self._products:
            return self._products[product_code].get(prop, "")
        return ""

    def ProductState(self, product_code: str) -> int:
        """Get product state."""
        return 5 if product_code in self._products else -1

    def ComponentClients(self, component_id: str) -> list[str]:
        """Get products owning a component."""
        return self._components.get(component_id, [])

    def ComponentPath(self, product_code: str, component_id: str) -> str:
        """Get component key path."""
        return self._component_paths.get((product_code, component_id), "")


@pytest.fixture
def mock_installer() -> MockInstaller:
    """Create a mock installer with sample Office products."""
    installer = MockInstaller()

    # Add Office products (ending with 0000000FF1CE)
    installer.add_product(
        "{90160000-000F-0000-1000-0000000FF1CE}",
        name="Microsoft Office Professional Plus 2016",
        version="16.0.4266.1001",
    )
    installer.add_product(
        "{90160000-0011-0000-0000-0000000FF1CE}",
        name="Microsoft Office Professional Plus 2016 - en-us",
        version="16.0.4266.1001",
    )

    # Add non-Office product
    installer.add_product(
        "{12345678-1234-1234-1234-123456789012}",
        name="Some Other Application",
        version="1.0.0",
    )

    # Add components
    installer.add_component(
        "{11111111-1111-1111-1111-111111111111}",
        clients=["{90160000-000F-0000-1000-0000000FF1CE}"],
        paths={
            "{90160000-000F-0000-1000-0000000FF1CE}": r"C:\Program Files\Office\WINWORD.EXE"
        },
    )
    installer.add_component(
        "{22222222-2222-2222-2222-222222222222}",
        clients=[
            "{90160000-000F-0000-1000-0000000FF1CE}",
            "{90160000-0011-0000-0000-0000000FF1CE}",
        ],
        paths={
            "{90160000-000F-0000-1000-0000000FF1CE}": "02:\\SOFTWARE\\Microsoft\\Office\\16.0"
        },
    )

    # Add orphaned component (no clients)
    installer.add_component(
        "{33333333-3333-3333-3333-333333333333}",
        clients=[],
    )

    return installer


class TestEnumerateProducts:
    """Tests for product enumeration."""

    def test_enumerate_all_products(self, mock_installer: MockInstaller) -> None:
        """Should enumerate all products."""
        products = list(enumerate_products(mock_installer))
        assert len(products) == 3

    def test_enumerate_office_only(self, mock_installer: MockInstaller) -> None:
        """Should filter to Office products only."""
        products = list(enumerate_products(mock_installer, office_only=True))
        assert len(products) == 2
        assert all(p.is_office for p in products)

    def test_product_info_populated(self, mock_installer: MockInstaller) -> None:
        """Product info should be populated."""
        products = list(enumerate_products(mock_installer, office_only=True))
        office_pro = next(
            (p for p in products if "Professional Plus" in p.name), None
        )
        assert office_pro is not None
        assert office_pro.name == "Microsoft Office Professional Plus 2016"
        assert office_pro.version == "16.0.4266.1001"
        assert office_pro.is_office is True

    def test_product_type_classification(self, mock_installer: MockInstaller) -> None:
        """Office products should have type classification."""
        products = list(enumerate_products(mock_installer, office_only=True))
        for p in products:
            assert p.product_type != ""  # Should have classification


class TestEnumerateComponents:
    """Tests for component enumeration."""

    def test_enumerate_all_components(self, mock_installer: MockInstaller) -> None:
        """Should enumerate all components."""
        components = list(enumerate_components(mock_installer))
        assert len(components) == 3


class TestComponentClients:
    """Tests for component client queries."""

    def test_get_clients_single(self, mock_installer: MockInstaller) -> None:
        """Should get single client."""
        clients = get_component_clients(
            "{11111111-1111-1111-1111-111111111111}",
            mock_installer,
        )
        assert len(clients) == 1
        assert "{90160000-000F-0000-1000-0000000FF1CE}" in clients

    def test_get_clients_multiple(self, mock_installer: MockInstaller) -> None:
        """Should get multiple clients."""
        clients = get_component_clients(
            "{22222222-2222-2222-2222-222222222222}",
            mock_installer,
        )
        assert len(clients) == 2

    def test_get_clients_orphaned(self, mock_installer: MockInstaller) -> None:
        """Orphaned component should have no clients."""
        clients = get_component_clients(
            "{33333333-3333-3333-3333-333333333333}",
            mock_installer,
        )
        assert len(clients) == 0


class TestComponentPath:
    """Tests for component path queries."""

    def test_get_file_path(self, mock_installer: MockInstaller) -> None:
        """Should get file path."""
        path = get_component_path(
            "{90160000-000F-0000-1000-0000000FF1CE}",
            "{11111111-1111-1111-1111-111111111111}",
            mock_installer,
        )
        assert path == r"C:\Program Files\Office\WINWORD.EXE"

    def test_get_registry_path(self, mock_installer: MockInstaller) -> None:
        """Should get registry path."""
        path = get_component_path(
            "{90160000-000F-0000-1000-0000000FF1CE}",
            "{22222222-2222-2222-2222-222222222222}",
            mock_installer,
        )
        assert path.startswith("02:")  # HKLM registry path

    def test_get_missing_path(self, mock_installer: MockInstaller) -> None:
        """Missing path should return empty string."""
        path = get_component_path(
            "{90160000-000F-0000-1000-0000000FF1CE}",
            "{33333333-3333-3333-3333-333333333333}",
            mock_installer,
        )
        assert path == ""


class TestIsOfficeComponent:
    """Tests for Office component detection."""

    def test_office_component(self, mock_installer: MockInstaller) -> None:
        """Component owned by Office should be detected."""
        assert is_office_component(
            "{11111111-1111-1111-1111-111111111111}",
            mock_installer,
        )

    def test_orphaned_component(self, mock_installer: MockInstaller) -> None:
        """Orphaned component with no clients."""
        # Orphaned components return False since they have no clients
        assert not is_office_component(
            "{33333333-3333-3333-3333-333333333333}",
            mock_installer,
        )


class TestMSIComponentScanner:
    """Tests for the MSIComponentScanner class."""

    def test_scan_office_products(self, mock_installer: MockInstaller) -> None:
        """Should scan and find Office products."""
        scanner = MSIComponentScanner()
        scanner._installer = mock_installer

        result = scanner.scan(office_only=True, include_components=True)

        assert len(result.office_products) == 2
        assert len(result.components) >= 2

    def test_scan_finds_orphaned(self, mock_installer: MockInstaller) -> None:
        """Should identify orphaned components."""
        scanner = MSIComponentScanner()
        scanner._installer = mock_installer

        result = scanner.scan(office_only=False, include_components=True)

        assert "{33333333-3333-3333-3333-333333333333}" in result.orphaned_components

    def test_scan_categorizes_paths(self, mock_installer: MockInstaller) -> None:
        """Should categorize file and registry paths."""
        scanner = MSIComponentScanner()
        scanner._installer = mock_installer

        result = scanner.scan(office_only=True, include_components=True)

        # Should have found the file path
        assert any("WINWORD.EXE" in p for p in result.file_paths)

        # Should have found the registry path (without the 02: prefix)
        assert any("Microsoft\\Office" in p for p in result.registry_paths)

    def test_scan_with_product_filter(self, mock_installer: MockInstaller) -> None:
        """Should filter to specific products."""
        scanner = MSIComponentScanner()
        scanner._installer = mock_installer

        result = scanner.scan(
            office_only=False,
            product_filter=["{90160000-000F-0000-1000-0000000FF1CE}"],
        )

        assert len(result.products) == 1
        assert result.products[0].product_code == "{90160000-000F-0000-1000-0000000FF1CE}"

    def test_write_scan_logs(
        self, mock_installer: MockInstaller, tmp_path: Path
    ) -> None:
        """Should write scan log files."""
        scanner = MSIComponentScanner()
        scanner._installer = mock_installer

        result = scanner.scan(office_only=True, include_components=True)
        files = scanner.write_scan_logs(result, tmp_path, prefix="test_")

        assert "files" in files
        assert "registry" in files
        assert "verbose" in files

        assert files["files"].exists()
        assert files["registry"].exists()
        assert files["verbose"].exists()

        # Check verbose file content
        content = files["verbose"].read_text()
        assert "Office Products: 2" in content


class TestScanResult:
    """Tests for ScanResult dataclass."""

    def test_default_values(self) -> None:
        """Default values should be empty lists/dicts."""
        result = ScanResult()
        assert result.products == []
        assert result.office_products == []
        assert result.components == {}
        assert result.orphaned_components == []
        assert result.file_paths == []
        assert result.registry_paths == []


class TestProductInfo:
    """Tests for ProductInfo dataclass."""

    def test_default_values(self) -> None:
        """Default values should be set correctly."""
        info = ProductInfo(product_code="{00000000-0000-0000-0000-000000000000}")
        assert info.name == ""
        assert info.version == ""
        assert info.install_state == MsiInstallState.UNKNOWN
        assert info.is_office is False


class TestComponentInfo:
    """Tests for ComponentInfo dataclass."""

    def test_default_values(self) -> None:
        """Default values should be set correctly."""
        info = ComponentInfo(component_id="{00000000-0000-0000-0000-000000000000}")
        assert info.key_path == ""
        assert info.clients == []
        assert info.state == MsiInstallState.UNKNOWN


class TestMsiInstallState:
    """Tests for MsiInstallState enum."""

    def test_state_values(self) -> None:
        """State values should match WI constants."""
        assert MsiInstallState.UNKNOWN == -1
        assert MsiInstallState.ABSENT == 2
        assert MsiInstallState.LOCAL == 3
        assert MsiInstallState.DEFAULT == 5


class TestWindowsInstallerError:
    """Tests for error handling."""

    def test_error_message(self) -> None:
        """WindowsInstallerError should include message."""
        error = WindowsInstallerError("test message")
        assert "test message" in str(error)
