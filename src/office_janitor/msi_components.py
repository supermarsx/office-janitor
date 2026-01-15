"""!
@brief Windows Installer component scanning and enumeration.
@details Provides a Python wrapper around the Windows Installer COM API
(``WindowsInstaller.Installer``) to enumerate products, components, and
their relationships. This is essential for thorough MSI cleanup as it
allows detection of orphaned components and validation of WI metadata.

The functionality mirrors the VBS ``ScanComponents`` function from
``OffScrub_O16msi.vbs`` and related helper functions that query WI state.

@note This module requires Windows and COM automation support. On non-Windows
platforms, stub classes are provided to allow module import without errors.
"""

from __future__ import annotations

import logging
import re
from collections.abc import Iterator, Sequence
from dataclasses import dataclass, field
from enum import IntEnum
from pathlib import Path
from typing import TYPE_CHECKING, Any, Final

from . import guid_utils

if TYPE_CHECKING:  # pragma: no cover - typing only
    pass

_LOGGER = logging.getLogger(__name__)


class MsiInstallState(IntEnum):
    """!
    @brief Windows Installer install state constants.
    @details Mirrors the ``msiInstallState*`` constants from the WI SDK.
    """

    NOT_USED = -7
    BAD_CONFIG = -6
    INCOMPLETE = -5
    SOURCE_ABSENT = -4
    MORE_DATA = -3
    INVALID_ARG = -2
    UNKNOWN = -1
    BROKEN = 0
    ADVERTISED = 1
    REMOVED = 1
    ABSENT = 2
    LOCAL = 3
    SOURCE = 4
    DEFAULT = 5


class MsiReinstallMode(IntEnum):
    """!
    @brief Windows Installer reinstall mode flags.
    """

    FILE_MISSING = 0x00000002
    FILE_OLDER_VERSION = 0x00000004
    FILE_EQUAL_VERSION = 0x00000008
    FILE_EXACT_VERSION = 0x00000010
    FILE_VERIFY = 0x00000020
    FILE_REPLACE = 0x00000040
    MACHINE_DATA = 0x00000080
    USER_DATA = 0x00000100
    SHORTCUT = 0x00000200
    PACKAGE = 0x00000400


# Office product code suffix pattern - ends with 0000000FF1CE
_OFFICE_SUFFIX_PATTERN: Final[re.Pattern[str]] = re.compile(r"0000000FF1CE\}?$", re.IGNORECASE)


@dataclass
class ProductInfo:
    """!
    @brief Information about an installed MSI product.
    """

    product_code: str
    """Standard GUID format."""

    name: str = ""
    """Product display name."""

    version: str = ""
    """Product version string."""

    install_location: str = ""
    """Installation directory."""

    install_source: str = ""
    """Original installation source path."""

    install_state: MsiInstallState = MsiInstallState.UNKNOWN
    """Current installation state."""

    package_code: str = ""
    """Package GUID identifying the specific MSI."""

    is_office: bool = False
    """True if product matches Office patterns."""

    product_type: str = ""
    """Office product type classification."""


@dataclass
class ComponentInfo:
    """!
    @brief Information about a Windows Installer component.
    """

    component_id: str
    """Component GUID in standard format."""

    key_path: str = ""
    """File or registry path that is the component's key resource."""

    clients: list[str] = field(default_factory=list)
    """List of product codes that own this component."""

    state: MsiInstallState = MsiInstallState.UNKNOWN
    """Component installation state for a specific product context."""


@dataclass
class ScanResult:
    """!
    @brief Results from a component scan operation.
    """

    products: list[ProductInfo] = field(default_factory=list)
    """All enumerated products."""

    office_products: list[ProductInfo] = field(default_factory=list)
    """Products identified as Office installations."""

    components: dict[str, ComponentInfo] = field(default_factory=dict)
    """Component ID to info mapping."""

    orphaned_components: list[str] = field(default_factory=list)
    """Components with no valid product clients."""

    file_paths: list[str] = field(default_factory=list)
    """All file key paths discovered."""

    registry_paths: list[str] = field(default_factory=list)
    """All registry key paths discovered."""


class WindowsInstallerError(RuntimeError):
    """!
    @brief Raised when Windows Installer operations fail.
    """


def _create_installer() -> Any:
    """!
    @brief Create a Windows Installer COM object.
    @return WindowsInstaller.Installer dispatch object.
    @throws WindowsInstallerError If COM creation fails.
    """
    try:
        import win32com.client

        return win32com.client.Dispatch("WindowsInstaller.Installer")
    except ImportError as e:
        raise WindowsInstallerError(
            "win32com is required for Windows Installer operations. "
            "Install with: pip install pywin32"
        ) from e
    except Exception as e:
        raise WindowsInstallerError(f"Failed to create Windows Installer object: {e}") from e


def _safe_product_info(installer: Any, product_code: str, property_name: str) -> str:
    """!
    @brief Safely retrieve a product property from Windows Installer.
    @param installer WI COM object.
    @param product_code Product GUID.
    @param property_name Property to retrieve.
    @return Property value or empty string on error.
    """
    try:
        return str(installer.ProductInfo(product_code, property_name) or "")
    except Exception:
        return ""


def enumerate_products(
    installer: Any | None = None,
    *,
    office_only: bool = False,
) -> Iterator[ProductInfo]:
    """!
    @brief Enumerate all installed MSI products.
    @param installer Optional pre-created WI COM object.
    @param office_only If True, yield only Office products.
    @yields ProductInfo for each installed product.

    @details Uses the WI ``ProductsEx`` or ``Products`` API to enumerate
    all registered MSI products on the system. This mirrors the VBS
    ``EnumProducts`` logic from the OffScrub scripts.
    """
    if installer is None:
        installer = _create_installer()

    try:
        # Try ProductsEx first (more complete on modern Windows)
        products = installer.ProductsEx("", "", 7)  # 7 = msiInstallContextAll
    except Exception:
        # Fall back to basic Products property
        try:
            products = installer.Products
        except Exception as e:
            _LOGGER.warning("Failed to enumerate products: %s", e)
            return

    for product_code in products:
        try:
            product_code = guid_utils.normalize_guid(str(product_code))
        except guid_utils.GuidError:
            continue

        is_office = bool(_OFFICE_SUFFIX_PATTERN.search(product_code))

        if office_only and not is_office:
            continue

        info = ProductInfo(
            product_code=product_code,
            name=_safe_product_info(installer, product_code, "ProductName"),
            version=_safe_product_info(installer, product_code, "VersionString"),
            install_location=_safe_product_info(installer, product_code, "InstallLocation"),
            install_source=_safe_product_info(installer, product_code, "InstallSource"),
            package_code=_safe_product_info(installer, product_code, "PackageCode"),
            is_office=is_office,
        )

        if is_office:
            info.product_type = guid_utils.classify_office_product(product_code)

        # Get install state
        try:
            state = installer.ProductState(product_code)
            info.install_state = MsiInstallState(state)
        except Exception:
            pass

        yield info


def enumerate_components(installer: Any | None = None) -> Iterator[str]:
    """!
    @brief Enumerate all registered Windows Installer components.
    @param installer Optional pre-created WI COM object.
    @yields Component GUIDs in standard format.

    @details This provides access to all WI components registered on the
    system, which is necessary for thorough cleanup of orphaned entries.
    """
    if installer is None:
        installer = _create_installer()

    try:
        components = installer.Components
        for comp_id in components:
            try:
                yield guid_utils.normalize_guid(str(comp_id))
            except guid_utils.GuidError:
                continue
    except Exception as e:
        _LOGGER.warning("Failed to enumerate components: %s", e)


def get_component_clients(component_id: str, installer: Any | None = None) -> list[str]:
    """!
    @brief Get all products that own a specific component.
    @param component_id Component GUID.
    @param installer Optional pre-created WI COM object.
    @return List of product codes that reference this component.

    @details A component with no clients is orphaned and can be cleaned up.
    """
    if installer is None:
        installer = _create_installer()

    clients: list[str] = []
    try:
        # Normalize component GUID
        comp_guid = guid_utils.normalize_guid(component_id)

        # Query clients
        client_list = installer.ComponentClients(comp_guid)
        for product_code in client_list:
            try:
                clients.append(guid_utils.normalize_guid(str(product_code)))
            except guid_utils.GuidError:
                clients.append(str(product_code))
    except Exception as e:
        _LOGGER.debug("Failed to get clients for component %s: %s", component_id, e)

    return clients


def get_component_path(product_code: str, component_id: str, installer: Any | None = None) -> str:
    """!
    @brief Get the key path for a component within a product context.
    @param product_code Product GUID providing context.
    @param component_id Component GUID.
    @param installer Optional pre-created WI COM object.
    @return Key path (file or registry) or empty string.

    @details The key path is the resource that Windows Installer uses to
    determine if a component is installed. It can be a file path or a
    registry key path (prefixed with registry hive indicators).
    """
    if installer is None:
        installer = _create_installer()

    try:
        path = installer.ComponentPath(product_code, component_id)
        return str(path) if path else ""
    except Exception:
        return ""


def get_component_state(
    product_code: str, component_id: str, installer: Any | None = None
) -> MsiInstallState:
    """!
    @brief Get the installation state of a component for a product.
    @param product_code Product GUID.
    @param component_id Component GUID.
    @param installer Optional pre-created WI COM object.
    @return Component installation state.
    """
    if installer is None:
        installer = _create_installer()

    try:
        # Open product to query feature/component state
        # This requires a more complex approach using MsiGetComponentState
        # For now, we infer state from the component path
        path = get_component_path(product_code, component_id, installer)
        if not path:
            return MsiInstallState.ABSENT
        if path.startswith("01:") or path.startswith("02:"):
            # Registry key path
            return MsiInstallState.LOCAL
        if Path(path).exists():
            return MsiInstallState.LOCAL
        return MsiInstallState.SOURCE_ABSENT
    except Exception:
        return MsiInstallState.UNKNOWN


def is_office_component(component_id: str, installer: Any | None = None) -> bool:
    """!
    @brief Check if a component belongs to any Office product.
    @param component_id Component GUID.
    @param installer Optional pre-created WI COM object.
    @return True if any client is an Office product.
    """
    clients = get_component_clients(component_id, installer)
    return any(_OFFICE_SUFFIX_PATTERN.search(pc) for pc in clients)


class MSIComponentScanner:
    """!
    @brief Scanner for Windows Installer components and products.

    @details Provides comprehensive scanning of WI metadata to identify
    Office products, their components, and orphaned registry entries.
    This mirrors the VBS ``ScanComponents`` function from OffScrub_O16msi.vbs.

    Usage:
    ```python
    scanner = MSIComponentScanner()
    result = scanner.scan(office_only=True)
    print(f"Found {len(result.office_products)} Office products")
    print(f"Found {len(result.orphaned_components)} orphaned components")
    ```
    """

    def __init__(self, logger: logging.Logger | None = None) -> None:
        """!
        @brief Initialize the scanner.
        @param logger Optional logger for diagnostic output.
        """
        self._logger = logger or _LOGGER
        self._installer: Any = None

    def _get_installer(self) -> Any:
        """Get or create the WI COM object."""
        if self._installer is None:
            self._installer = _create_installer()
        return self._installer

    def scan(
        self,
        *,
        office_only: bool = True,
        include_components: bool = True,
        product_filter: Sequence[str] | None = None,
    ) -> ScanResult:
        """!
        @brief Perform a comprehensive scan of WI products and components.
        @param office_only If True, focus only on Office products.
        @param include_components If True, enumerate components for products.
        @param product_filter Optional list of product codes to limit scanning.
        @return ScanResult with all discovered information.
        """
        result = ScanResult()
        installer = self._get_installer()

        self._logger.info("Starting Windows Installer scan")

        # Enumerate products
        for product in enumerate_products(installer, office_only=office_only):
            if product_filter and product.product_code not in product_filter:
                continue

            result.products.append(product)
            if product.is_office:
                result.office_products.append(product)

            self._logger.debug(
                "Found product: %s (%s)",
                product.name or product.product_code,
                product.product_type or "Unknown type",
            )

        self._logger.info(
            "Enumerated %d products (%d Office)",
            len(result.products),
            len(result.office_products),
        )

        if not include_components:
            return result

        # Enumerate components for Office products
        office_product_codes = {p.product_code for p in result.office_products}

        for comp_id in enumerate_components(installer):
            clients = get_component_clients(comp_id, installer)

            # Check if this component is owned by any Office product
            office_clients = [c for c in clients if c in office_product_codes]

            if not office_clients:
                # Check if it matches Office pattern even without known client
                if office_only and not any(_OFFICE_SUFFIX_PATTERN.search(c) for c in clients):
                    continue

            # Get component info using first available client
            client_for_path = (
                office_clients[0] if office_clients else (clients[0] if clients else "")
            )

            key_path = ""
            if client_for_path:
                key_path = get_component_path(client_for_path, comp_id, installer)

            comp_info = ComponentInfo(
                component_id=comp_id,
                key_path=key_path,
                clients=clients,
            )
            result.components[comp_id] = comp_info

            # Categorize the key path
            if key_path:
                if key_path.startswith("01:") or key_path.startswith("02:"):
                    # Registry path (01: = HKCU, 02: = HKLM)
                    result.registry_paths.append(key_path[3:])
                elif not key_path.startswith("00:"):
                    # File path
                    result.file_paths.append(key_path)

            # Check for orphaned components
            if not clients:
                result.orphaned_components.append(comp_id)

        self._logger.info(
            "Enumerated %d components, %d file paths, %d registry paths, %d orphaned",
            len(result.components),
            len(result.file_paths),
            len(result.registry_paths),
            len(result.orphaned_components),
        )

        return result

    def scan_for_product(self, product_code: str) -> ScanResult:
        """!
        @brief Scan components for a specific product.
        @param product_code Product GUID to scan.
        @return ScanResult focused on the specified product.
        """
        return self.scan(
            office_only=False,
            include_components=True,
            product_filter=[guid_utils.normalize_guid(product_code)],
        )

    def get_product_components(self, product_code: str) -> list[ComponentInfo]:
        """!
        @brief Get all components belonging to a specific product.
        @param product_code Product GUID.
        @return List of ComponentInfo for the product's components.
        """
        installer = self._get_installer()
        product_code = guid_utils.normalize_guid(product_code)

        components: list[ComponentInfo] = []

        for comp_id in enumerate_components(installer):
            clients = get_component_clients(comp_id, installer)
            if product_code in clients:
                key_path = get_component_path(product_code, comp_id, installer)
                components.append(
                    ComponentInfo(
                        component_id=comp_id,
                        key_path=key_path,
                        clients=clients,
                        state=get_component_state(product_code, comp_id, installer),
                    )
                )

        return components

    def write_scan_logs(
        self,
        result: ScanResult,
        output_dir: Path,
        *,
        prefix: str = "",
    ) -> dict[str, Path]:
        """!
        @brief Write scan results to log files.
        @param result ScanResult to write.
        @param output_dir Directory for output files.
        @param prefix Optional prefix for filenames.
        @return Dictionary mapping log type to file path.

        @details Generates files similar to VBS OffScrub output:
        - FileList.txt: All file paths from component key paths
        - RegList.txt: All registry paths from component key paths
        - CompVerbose.txt: Detailed component information
        """
        output_dir.mkdir(parents=True, exist_ok=True)
        files: dict[str, Path] = {}

        # FileList.txt
        file_list_path = output_dir / f"{prefix}FileList.txt"
        with file_list_path.open("w", encoding="utf-8") as f:
            for path in sorted(set(result.file_paths)):
                f.write(f"{path}\n")
        files["files"] = file_list_path

        # RegList.txt
        reg_list_path = output_dir / f"{prefix}RegList.txt"
        with reg_list_path.open("w", encoding="utf-8") as f:
            for path in sorted(set(result.registry_paths)):
                f.write(f"{path}\n")
        files["registry"] = reg_list_path

        # CompVerbose.txt
        comp_verbose_path = output_dir / f"{prefix}CompVerbose.txt"
        with comp_verbose_path.open("w", encoding="utf-8") as f:
            f.write("=" * 80 + "\n")
            f.write("Office Janitor - Component Scan Results\n")
            f.write("=" * 80 + "\n\n")

            f.write(f"Products: {len(result.products)}\n")
            f.write(f"Office Products: {len(result.office_products)}\n")
            f.write(f"Components: {len(result.components)}\n")
            f.write(f"Orphaned Components: {len(result.orphaned_components)}\n\n")

            f.write("-" * 80 + "\n")
            f.write("PRODUCTS\n")
            f.write("-" * 80 + "\n")
            for p in result.office_products:
                f.write(f"\n{p.product_code}\n")
                f.write(f"  Name: {p.name}\n")
                f.write(f"  Type: {p.product_type}\n")
                f.write(f"  Version: {p.version}\n")
                f.write(f"  Location: {p.install_location}\n")
                f.write(f"  State: {p.install_state.name}\n")

            if result.orphaned_components:
                f.write("\n" + "-" * 80 + "\n")
                f.write("ORPHANED COMPONENTS\n")
                f.write("-" * 80 + "\n")
                for comp_id in result.orphaned_components:
                    f.write(f"{comp_id}\n")

        files["verbose"] = comp_verbose_path

        self._logger.info(
            "Wrote scan logs to %s: %s",
            output_dir,
            ", ".join(files.keys()),
        )

        return files


def scan_office_products(*, logger: logging.Logger | None = None) -> ScanResult:
    """!
    @brief Convenience function to scan for Office products and components.
    @param logger Optional logger.
    @return ScanResult with Office products and their components.
    """
    scanner = MSIComponentScanner(logger=logger)
    return scanner.scan(office_only=True, include_components=True)


def list_office_products(*, logger: logging.Logger | None = None) -> list[ProductInfo]:
    """!
    @brief List all installed Office MSI products.
    @param logger Optional logger.
    @return List of ProductInfo for Office products.
    """
    scanner = MSIComponentScanner(logger=logger)
    result = scanner.scan(office_only=True, include_components=False)
    return result.office_products


__all__ = [
    "ComponentInfo",
    "MSIComponentScanner",
    "MsiInstallState",
    "MsiReinstallMode",
    "ProductInfo",
    "ScanResult",
    "WindowsInstallerError",
    "enumerate_components",
    "enumerate_products",
    "get_component_clients",
    "get_component_path",
    "get_component_state",
    "is_office_component",
    "list_office_products",
    "scan_office_products",
]
