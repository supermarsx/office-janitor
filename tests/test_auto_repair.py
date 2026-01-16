"""!
@brief Tests for auto_repair module.
@details Verifies auto-repair detection, planning, and execution logic.
"""

from __future__ import annotations

from unittest.mock import MagicMock, patch

from office_janitor import auto_repair


class TestRepairMethodEnum:
    """Tests for RepairMethod enumeration."""

    def test_repair_method_values(self) -> None:
        """RepairMethod should have expected values."""
        assert auto_repair.RepairMethod.C2R.value == "c2r"
        assert auto_repair.RepairMethod.ODT.value == "odt"
        assert auto_repair.RepairMethod.MSI.value == "msi"
        assert auto_repair.RepairMethod.AUTO.value == "auto"


class TestRepairStrategyEnum:
    """Tests for RepairStrategy enumeration."""

    def test_repair_strategy_values(self) -> None:
        """RepairStrategy should have expected values."""
        assert auto_repair.RepairStrategy.QUICK.value == "quick"
        assert auto_repair.RepairStrategy.FULL.value == "full"
        assert auto_repair.RepairStrategy.INCREMENTAL.value == "incremental"
        assert auto_repair.RepairStrategy.REINSTALL.value == "reinstall"


class TestDetectedOfficeProduct:
    """Tests for DetectedOfficeProduct dataclass."""

    def test_product_creation(self) -> None:
        """DetectedOfficeProduct should be created with required fields."""
        product = auto_repair.DetectedOfficeProduct(
            product_id="O365ProPlusRetail",
            product_name="Microsoft 365 Apps",
            version="16.0.18324.20000",
            install_type="c2r",
            platform="x64",
            culture="en-us",
        )
        assert product.product_id == "O365ProPlusRetail"
        assert product.product_name == "Microsoft 365 Apps"
        assert product.install_type == "c2r"
        assert product.can_repair is True
        assert product.repair_methods == []

    def test_product_with_optional_fields(self) -> None:
        """DetectedOfficeProduct should accept optional fields."""
        product = auto_repair.DetectedOfficeProduct(
            product_id="Office2019",
            product_name="Office 2019",
            version="16.0.10000",
            install_type="msi",
            platform="x86",
            culture="de-de",
            product_code="{GUID-HERE}",
            repair_methods=[auto_repair.RepairMethod.MSI],
        )
        assert product.product_code == "{GUID-HERE}"
        assert auto_repair.RepairMethod.MSI in product.repair_methods


class TestAutoRepairPlan:
    """Tests for AutoRepairPlan dataclass."""

    def test_plan_creation(self) -> None:
        """AutoRepairPlan should be created correctly."""
        products = [
            auto_repair.DetectedOfficeProduct(
                product_id="test",
                product_name="Test Product",
                version="1.0",
                install_type="c2r",
                platform="x64",
                culture="en-us",
            )
        ]
        plan = auto_repair.AutoRepairPlan(
            products=products,
            recommended_method=auto_repair.RepairMethod.C2R,
            recommended_strategy=auto_repair.RepairStrategy.QUICK,
        )
        assert len(plan.products) == 1
        assert plan.recommended_method == auto_repair.RepairMethod.C2R
        assert plan.requires_internet is False
        assert plan.estimated_time_minutes == 15

    def test_plan_with_full_strategy(self) -> None:
        """Plan with FULL strategy should indicate internet requirement."""
        plan = auto_repair.AutoRepairPlan(
            products=[],
            recommended_method=auto_repair.RepairMethod.C2R,
            recommended_strategy=auto_repair.RepairStrategy.FULL,
            requires_internet=True,
        )
        assert plan.requires_internet is True


class TestAutoRepairResult:
    """Tests for AutoRepairResult dataclass."""

    def test_successful_result(self) -> None:
        """Successful result should have correct summary."""
        result = auto_repair.AutoRepairResult(
            success=True,
            products_repaired=["Office 365"],
            products_failed=[],
            products_skipped=[],
            total_duration=120.5,
            method_used=auto_repair.RepairMethod.C2R,
        )
        assert result.success is True
        assert "1 repaired" in result.summary
        assert "120.5s" in result.summary

    def test_failed_result(self) -> None:
        """Failed result should have correct summary."""
        result = auto_repair.AutoRepairResult(
            success=False,
            products_repaired=["Product A"],
            products_failed=["Product B"],
            products_skipped=[],
            total_duration=60.0,
            method_used=auto_repair.RepairMethod.C2R,
            errors=["Repair failed"],
        )
        assert result.success is False
        assert "1 repaired" in result.summary
        assert "1 failed" in result.summary


class TestProductNameMapping:
    """Tests for product name lookup."""

    def test_get_known_product_name(self) -> None:
        """Known product IDs should return display names."""
        name = auto_repair._get_product_display_name("O365ProPlusRetail")
        assert name == "Microsoft 365 Apps for enterprise"

    def test_get_unknown_product_name(self) -> None:
        """Unknown product IDs should return the ID itself."""
        name = auto_repair._get_product_display_name("UnknownProduct")
        assert name == "UnknownProduct"


class TestCreateRepairPlan:
    """Tests for create_repair_plan function."""

    def test_plan_with_no_products(self) -> None:
        """Plan with no products should still be created."""
        plan = auto_repair.create_repair_plan(products=[])
        assert plan.products == []
        assert plan.recommended_method in auto_repair.RepairMethod

    def test_plan_with_c2r_products(self) -> None:
        """Plan with C2R products should recommend C2R method."""
        products = [
            auto_repair.DetectedOfficeProduct(
                product_id="O365ProPlusRetail",
                product_name="Microsoft 365",
                version="16.0",
                install_type="c2r",
                platform="x64",
                culture="en-us",
            )
        ]
        plan = auto_repair.create_repair_plan(products=products)
        assert plan.recommended_method == auto_repair.RepairMethod.C2R

    def test_plan_with_msi_products(self) -> None:
        """Plan with MSI products should recommend MSI method."""
        products = [
            auto_repair.DetectedOfficeProduct(
                product_id="{GUID}",
                product_name="Office 2019",
                version="16.0",
                install_type="msi",
                platform="x64",
                culture="en-us",
            )
        ]
        plan = auto_repair.create_repair_plan(products=products)
        assert plan.recommended_method == auto_repair.RepairMethod.MSI

    def test_plan_full_strategy_adds_warnings(self) -> None:
        """Full strategy should add internet warning."""
        plan = auto_repair.create_repair_plan(
            products=[],
            strategy=auto_repair.RepairStrategy.FULL,
        )
        assert any("internet" in w.lower() for w in plan.warnings)


class TestDetectionFunctions:
    """Tests for detection functions."""

    @patch("office_janitor.auto_repair.registry_tools")
    def test_detect_c2r_products_no_installation(self, mock_registry: MagicMock) -> None:
        """No C2R installation should return empty list."""
        mock_registry.get_value.return_value = None
        products = auto_repair._detect_c2r_products()
        assert products == []

    @patch("office_janitor.auto_repair.registry_tools")
    def test_detect_c2r_products_with_installation(self, mock_registry: MagicMock) -> None:
        """C2R installation should be detected."""
        mock_registry.get_value.side_effect = [
            "x64",  # Platform
            "16.0.18324.20000",  # Version
            "en-us",  # Culture
            "O365ProPlusRetail",  # ProductReleaseIds
            "http://cdn.office.net",  # CDNBaseUrl
        ]
        products = auto_repair._detect_c2r_products()
        assert len(products) == 1
        assert products[0].product_id == "O365ProPlusRetail"
        assert products[0].install_type == "c2r"


class TestConvenienceFunctions:
    """Tests for convenience functions."""

    @patch("office_janitor.auto_repair.execute_auto_repair")
    def test_quick_auto_repair(self, mock_execute: MagicMock) -> None:
        """quick_auto_repair should call execute_auto_repair with QUICK strategy."""
        mock_execute.return_value = auto_repair.AutoRepairResult(
            success=True,
            products_repaired=[],
            products_failed=[],
            products_skipped=[],
            total_duration=0,
            method_used=auto_repair.RepairMethod.AUTO,
        )
        auto_repair.quick_auto_repair(dry_run=True)
        mock_execute.assert_called_once()
        call_kwargs = mock_execute.call_args.kwargs
        assert call_kwargs["strategy"] == auto_repair.RepairStrategy.QUICK
        assert call_kwargs["dry_run"] is True

    @patch("office_janitor.auto_repair.execute_auto_repair")
    def test_full_auto_repair(self, mock_execute: MagicMock) -> None:
        """full_auto_repair should call execute_auto_repair with FULL strategy."""
        mock_execute.return_value = auto_repair.AutoRepairResult(
            success=True,
            products_repaired=[],
            products_failed=[],
            products_skipped=[],
            total_duration=0,
            method_used=auto_repair.RepairMethod.AUTO,
        )
        auto_repair.full_auto_repair(dry_run=True)
        mock_execute.assert_called_once()
        call_kwargs = mock_execute.call_args.kwargs
        assert call_kwargs["strategy"] == auto_repair.RepairStrategy.FULL

    @patch("office_janitor.auto_repair.repair_module.quick_repair")
    def test_repair_c2r_quick(self, mock_repair: MagicMock) -> None:
        """repair_c2r_quick should call repair module."""
        auto_repair.repair_c2r_quick(culture="de-de", dry_run=True)
        mock_repair.assert_called_once_with(culture="de-de", dry_run=True)

    @patch("office_janitor.auto_repair.repair_module.full_repair")
    def test_repair_c2r_full(self, mock_repair: MagicMock) -> None:
        """repair_c2r_full should call repair module."""
        auto_repair.repair_c2r_full(culture="fr-fr", dry_run=True)
        mock_repair.assert_called_once_with(culture="fr-fr", dry_run=True)
