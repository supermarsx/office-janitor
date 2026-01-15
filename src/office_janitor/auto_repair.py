"""!
@brief Auto-repair orchestration for Office installations.
@details Provides intelligent auto-detection and repair capabilities for both
MSI and Click-to-Run Office installations. Supports multiple repair strategies
and automatic fallback mechanisms.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import cast

from . import command_runner, constants, logging_ext, registry_tools
from . import repair as repair_module
from .exec_utils import CommandResult

# ---------------------------------------------------------------------------
# Constants and Enumerations
# ---------------------------------------------------------------------------


class RepairMethod(Enum):
    """!
    @brief Available repair methods for Office installations.
    """

    C2R = "c2r"  # OfficeClickToRun.exe
    ODT = "odt"  # Office Deployment Tool setup.exe
    MSI = "msi"  # Windows Installer repair
    AUTO = "auto"  # Auto-detect best method


class RepairStrategy(Enum):
    """!
    @brief Repair strategies for different scenarios.
    """

    QUICK = "quick"  # Fast local repair
    FULL = "full"  # Complete online repair
    INCREMENTAL = "incremental"  # Repair only failed components
    REINSTALL = "reinstall"  # Full reinstall


# ---------------------------------------------------------------------------
# Data Classes
# ---------------------------------------------------------------------------


@dataclass
class DetectedOfficeProduct:
    """!
    @brief Represents a detected Office product for repair.
    """

    product_id: str
    product_name: str
    version: str
    install_type: str  # 'c2r', 'msi', 'appx'
    platform: str  # 'x86', 'x64'
    culture: str
    install_path: Path | None = None
    product_code: str | None = None  # For MSI
    release_id: str | None = None  # For C2R
    channel: str | None = None  # For C2R
    can_repair: bool = True
    repair_methods: list[RepairMethod] = field(default_factory=list)


@dataclass
class AutoRepairPlan:
    """!
    @brief Plan for auto-repair operations.
    """

    products: list[DetectedOfficeProduct]
    recommended_method: RepairMethod
    recommended_strategy: RepairStrategy
    warnings: list[str] = field(default_factory=list)
    requires_internet: bool = False
    estimated_time_minutes: int = 15
    dry_run: bool = False


@dataclass
class AutoRepairResult:
    """!
    @brief Result of an auto-repair operation.
    """

    success: bool
    products_repaired: list[str]
    products_failed: list[str]
    products_skipped: list[str]
    total_duration: float
    method_used: RepairMethod
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    details: dict[str, object] = field(default_factory=dict)

    @property
    def summary(self) -> str:
        """!
        @brief Human-readable summary of the repair result.
        """
        if self.success:
            return (
                f"Auto-repair completed: {len(self.products_repaired)} repaired, "
                f"{len(self.products_skipped)} skipped in {self.total_duration:.1f}s"
            )
        return (
            f"Auto-repair partially failed: {len(self.products_repaired)} repaired, "
            f"{len(self.products_failed)} failed, {len(self.products_skipped)} skipped"
        )


# ---------------------------------------------------------------------------
# Detection Functions
# ---------------------------------------------------------------------------


def detect_office_products() -> list[DetectedOfficeProduct]:
    """!
    @brief Detect all installed Office products for repair.
    @returns List of DetectedOfficeProduct instances.
    """
    log = logging_ext.get_human_logger()
    products: list[DetectedOfficeProduct] = []

    # Detect C2R installations
    c2r_products = _detect_c2r_products()
    products.extend(c2r_products)

    # Detect MSI installations
    msi_products = _detect_msi_products()
    products.extend(msi_products)

    log.info(f"Detected {len(products)} Office product(s) for repair")
    return products


def _detect_c2r_products() -> list[DetectedOfficeProduct]:
    """!
    @brief Detect Click-to-Run Office products.
    """
    products: list[DetectedOfficeProduct] = []

    c2r_config_path = r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    try:
        # Check if C2R is installed
        platform = registry_tools.get_value(constants.HKLM, c2r_config_path, "Platform")
        if not platform:
            return products

        version = registry_tools.get_value(constants.HKLM, c2r_config_path, "VersionToReport")
        culture = registry_tools.get_value(constants.HKLM, c2r_config_path, "ClientCulture")
        product_ids = registry_tools.get_value(constants.HKLM, c2r_config_path, "ProductReleaseIds")
        channel = registry_tools.get_value(constants.HKLM, c2r_config_path, "CDNBaseUrl")

        # Parse product release IDs
        if product_ids:
            for release_id in str(product_ids).split(","):
                release_id = release_id.strip()
                if not release_id:
                    continue

                product = DetectedOfficeProduct(
                    product_id=release_id,
                    product_name=_get_product_display_name(release_id),
                    version=str(version) if version else "unknown",
                    install_type="c2r",
                    platform="x64" if "x64" in str(platform).lower() else "x86",
                    culture=str(culture).lower() if culture else "en-us",
                    release_id=release_id,
                    channel=str(channel) if channel else None,
                    repair_methods=[RepairMethod.C2R, RepairMethod.ODT],
                )
                products.append(product)

    except Exception as e:
        logging_ext.get_human_logger().debug(f"Error detecting C2R products: {e}")

    return products


def _detect_msi_products() -> list[DetectedOfficeProduct]:
    """!
    @brief Detect MSI-based Office products.
    """
    products: list[DetectedOfficeProduct] = []

    # Check common MSI Office registry paths
    uninstall_paths = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
    ]

    for base_path in uninstall_paths:
        try:
            subkeys = list(registry_tools.iter_subkeys(constants.HKLM, base_path))
            if not subkeys:
                continue

            for subkey in subkeys:
                key_path = f"{base_path}\\{subkey}"
                display_name = registry_tools.get_value(constants.HKLM, key_path, "DisplayName")

                if not display_name:
                    continue

                # Check if this is an Office product
                name_lower = str(display_name).lower()
                if not any(
                    x in name_lower
                    for x in [
                        "microsoft office",
                        "microsoft 365",
                        "microsoft word",
                        "microsoft excel",
                    ]
                ):
                    continue

                # Skip C2R entries (handled separately)
                install_source = registry_tools.get_value(constants.HKLM, key_path, "InstallSource")
                if install_source and "clicktorun" in str(install_source).lower():
                    continue

                version = registry_tools.get_value(constants.HKLM, key_path, "DisplayVersion")
                install_location = registry_tools.get_value(
                    constants.HKLM, key_path, "InstallLocation"
                )

                # Determine platform from path
                platform = "x86" if "WOW6432Node" in base_path else "x64"

                product = DetectedOfficeProduct(
                    product_id=subkey,
                    product_name=str(display_name),
                    version=str(version) if version else "unknown",
                    install_type="msi",
                    platform=platform,
                    culture="en-us",  # MSI doesn't store culture easily
                    install_path=Path(str(install_location)) if install_location else None,
                    product_code=subkey if subkey.startswith("{") else None,
                    repair_methods=[RepairMethod.MSI],
                )
                products.append(product)

        except Exception as e:
            logging_ext.get_human_logger().debug(f"Error detecting MSI products: {e}")

    return products


def _get_product_display_name(release_id: str) -> str:
    """!
    @brief Get a human-readable display name for a release ID.
    """
    # Common C2R product IDs to display names
    PRODUCT_NAMES = {
        "O365ProPlusRetail": "Microsoft 365 Apps for enterprise",
        "O365BusinessRetail": "Microsoft 365 Apps for business",
        "O365HomePremRetail": "Microsoft 365 Personal/Family",
        "ProPlus2024Volume": "Office LTSC 2024 Professional Plus",
        "ProPlus2021Volume": "Office LTSC 2021 Professional Plus",
        "ProPlus2019Volume": "Office 2019 Professional Plus",
        "VisioProRetail": "Visio Professional",
        "VisioProVolume": "Visio Professional (Volume)",
        "ProjectProRetail": "Project Professional",
        "ProjectProVolume": "Project Professional (Volume)",
        "AccessRetail": "Microsoft Access",
        "ExcelRetail": "Microsoft Excel",
        "WordRetail": "Microsoft Word",
        "PowerPointRetail": "Microsoft PowerPoint",
        "OutlookRetail": "Microsoft Outlook",
        "PublisherRetail": "Microsoft Publisher",
        "OneNoteRetail": "Microsoft OneNote",
    }

    return PRODUCT_NAMES.get(release_id, release_id)


# ---------------------------------------------------------------------------
# Planning Functions
# ---------------------------------------------------------------------------


def create_repair_plan(
    products: list[DetectedOfficeProduct] | None = None,
    *,
    method: RepairMethod = RepairMethod.AUTO,
    strategy: RepairStrategy = RepairStrategy.QUICK,
    dry_run: bool = False,
) -> AutoRepairPlan:
    """!
    @brief Create a repair plan for detected Office products.
    @param products List of products to repair (auto-detect if None).
    @param method Preferred repair method.
    @param strategy Repair strategy to use.
    @param dry_run Whether this is a dry-run.
    @returns AutoRepairPlan instance.
    """
    if products is None:
        products = detect_office_products()

    warnings: list[str] = []
    requires_internet = strategy == RepairStrategy.FULL

    # Determine recommended method
    recommended_method = method
    if method == RepairMethod.AUTO:
        # Prefer C2R for C2R installations, MSI for MSI
        c2r_count = sum(1 for p in products if p.install_type == "c2r")
        msi_count = sum(1 for p in products if p.install_type == "msi")

        if c2r_count > msi_count:
            recommended_method = RepairMethod.C2R
        elif msi_count > 0:
            recommended_method = RepairMethod.MSI
        else:
            recommended_method = RepairMethod.C2R  # Default

    # Estimate time
    base_time = 10  # minutes
    if strategy == RepairStrategy.FULL:
        base_time = 45
    elif strategy == RepairStrategy.REINSTALL:
        base_time = 60

    estimated_time = base_time + (len(products) * 5)

    # Add warnings
    if strategy == RepairStrategy.FULL:
        warnings.append("Full repair may reinstall previously excluded applications")
        warnings.append("Internet connectivity required")

    if any(p.install_type == "msi" for p in products):
        warnings.append("MSI repairs use Windows Installer repair mechanism")

    return AutoRepairPlan(
        products=products,
        recommended_method=recommended_method,
        recommended_strategy=strategy,
        warnings=warnings,
        requires_internet=requires_internet,
        estimated_time_minutes=estimated_time,
        dry_run=dry_run,
    )


# ---------------------------------------------------------------------------
# Repair Execution Functions
# ---------------------------------------------------------------------------


def execute_auto_repair(
    plan: AutoRepairPlan | None = None,
    *,
    method: RepairMethod = RepairMethod.AUTO,
    strategy: RepairStrategy = RepairStrategy.QUICK,
    culture: str | None = None,
    platform: str | None = None,
    silent: bool = True,
    dry_run: bool = False,
    force: bool = False,
) -> AutoRepairResult:
    """!
    @brief Execute auto-repair for Office installations.
    @param plan Optional pre-created repair plan.
    @param method Repair method to use.
    @param strategy Repair strategy.
    @param culture Language/culture code.
    @param platform Architecture (x86/x64).
    @param silent Run silently without UI.
    @param dry_run Simulate without executing.
    @param force Skip confirmations.
    @returns AutoRepairResult with repair outcome.
    """
    import time

    log = logging_ext.get_human_logger()
    mlog = logging_ext.get_machine_logger()
    start_time = time.perf_counter()

    # Create plan if not provided
    if plan is None:
        plan = create_repair_plan(method=method, strategy=strategy, dry_run=dry_run)

    mlog.info(
        "auto_repair_start",
        extra={
            "event": "auto_repair_start",
            "product_count": len(plan.products),
            "method": plan.recommended_method.value,
            "strategy": plan.recommended_strategy.value,
            "dry_run": dry_run,
        },
    )

    if not plan.products:
        log.warning("No Office products detected for repair")
        return AutoRepairResult(
            success=True,
            products_repaired=[],
            products_failed=[],
            products_skipped=[],
            total_duration=time.perf_counter() - start_time,
            method_used=plan.recommended_method,
            warnings=["No Office products detected"],
        )

    # Log detected products
    log.info(f"Found {len(plan.products)} Office product(s) to repair:")
    for product in plan.products:
        log.info(f"  - {product.product_name} ({product.version}, {product.install_type})")

    # Execute repair based on method
    products_repaired: list[str] = []
    products_failed: list[str] = []
    products_skipped: list[str] = []
    errors: list[str] = []
    details: dict[str, object] = {}

    if plan.recommended_method in (RepairMethod.C2R, RepairMethod.AUTO):
        # Use C2R repair for Click-to-Run products
        c2r_products = [p for p in plan.products if p.install_type == "c2r"]
        if c2r_products:
            c2r_result = _repair_c2r_products(
                c2r_products,
                strategy=plan.recommended_strategy,
                culture=culture,
                platform=platform,
                silent=silent,
                dry_run=dry_run,
            )
            products_repaired.extend(cast(list[str], c2r_result.get("repaired", [])))
            products_failed.extend(cast(list[str], c2r_result.get("failed", [])))
            errors.extend(cast(list[str], c2r_result.get("errors", [])))
            details["c2r"] = c2r_result

    if plan.recommended_method in (RepairMethod.MSI, RepairMethod.AUTO):
        # Use MSI repair for MSI products
        msi_products = [p for p in plan.products if p.install_type == "msi"]
        if msi_products:
            msi_result = _repair_msi_products(
                msi_products,
                dry_run=dry_run,
            )
            products_repaired.extend(cast(list[str], msi_result.get("repaired", [])))
            products_failed.extend(cast(list[str], msi_result.get("failed", [])))
            errors.extend(cast(list[str], msi_result.get("errors", [])))
            details["msi"] = msi_result

    if plan.recommended_method == RepairMethod.ODT:
        # Use ODT for repair
        odt_result = _repair_via_odt(
            plan.products,
            strategy=plan.recommended_strategy,
            dry_run=dry_run,
        )
        products_repaired.extend(cast(list[str], odt_result.get("repaired", [])))
        products_failed.extend(cast(list[str], odt_result.get("failed", [])))
        errors.extend(cast(list[str], odt_result.get("errors", [])))
        details["odt"] = odt_result

    # Mark products with unsupported install types as skipped
    for product in plan.products:
        is_repaired = product.product_name in products_repaired
        is_failed = product.product_name in products_failed
        if not is_repaired and not is_failed:
            if product.install_type not in ("c2r", "msi"):
                products_skipped.append(product.product_name)

    duration = time.perf_counter() - start_time
    success = len(products_failed) == 0 and len(products_repaired) > 0

    result = AutoRepairResult(
        success=success,
        products_repaired=products_repaired,
        products_failed=products_failed,
        products_skipped=products_skipped,
        total_duration=duration,
        method_used=plan.recommended_method,
        errors=errors,
        warnings=plan.warnings,
        details=details,
    )

    mlog.info(
        "auto_repair_complete",
        extra={
            "event": "auto_repair_complete",
            "success": success,
            "repaired_count": len(products_repaired),
            "failed_count": len(products_failed),
            "duration": duration,
        },
    )

    if success:
        log.info(result.summary)
    else:
        log.error(result.summary)
        for error in errors:
            log.error(f"  - {error}")

    return result


def _repair_c2r_products(
    products: list[DetectedOfficeProduct],
    *,
    strategy: RepairStrategy,
    culture: str | None = None,
    platform: str | None = None,
    silent: bool = True,
    dry_run: bool = False,
) -> dict[str, object]:
    """!
    @brief Repair C2R products using OfficeClickToRun.exe.
    """
    log = logging_ext.get_human_logger()
    repaired: list[str] = []
    failed: list[str] = []
    errors: list[str] = []

    # C2R repairs all products at once, so we only need one repair call
    if not products:
        return {"repaired": repaired, "failed": failed, "errors": errors}

    # Use the first product's attributes as reference
    ref_product = products[0]
    effective_culture = culture or ref_product.culture
    effective_platform = platform or ref_product.platform

    log.info(f"Repairing {len(products)} C2R product(s)...")

    # Determine repair type based on strategy
    if strategy == RepairStrategy.FULL:
        config = repair_module.RepairConfig.full_repair(
            platform=effective_platform,
            culture=effective_culture,
            silent=silent,
        )
    else:
        config = repair_module.RepairConfig.quick_repair(
            platform=effective_platform,
            culture=effective_culture,
            silent=silent,
        )

    result = repair_module.run_repair(config, dry_run=dry_run)

    if result.success or result.skipped:
        repaired.extend(p.product_name for p in products)
    else:
        failed.extend(p.product_name for p in products)
        errors.append(f"C2R repair failed: {result.error_message}")

    return {
        "repaired": repaired,
        "failed": failed,
        "errors": errors,
        "result": {
            "success": result.success,
            "return_code": result.return_code,
            "duration": result.duration,
        },
    }


def _repair_msi_products(
    products: list[DetectedOfficeProduct],
    *,
    dry_run: bool = False,
) -> dict[str, object]:
    """!
    @brief Repair MSI products using Windows Installer.
    """
    log = logging_ext.get_human_logger()
    repaired: list[str] = []
    failed: list[str] = []
    errors: list[str] = []

    for product in products:
        if not product.product_code:
            log.warning(f"No product code for {product.product_name}, skipping MSI repair")
            continue

        log.info(f"Repairing MSI product: {product.product_name}")

        if dry_run:
            log.info(f"[DRY-RUN] Would repair MSI: {product.product_code}")
            repaired.append(product.product_name)
            continue

        # Run msiexec /fa (full reinstall with all files)
        command = [
            "msiexec.exe",
            "/fa",
            product.product_code,
            "/qn",  # Quiet, no UI
            "REBOOT=ReallySuppress",
        ]

        result = command_runner.run_command(
            command,
            event="msi_repair",
            timeout=1800,  # 30 minutes
        )

        if result.returncode == 0:
            repaired.append(product.product_name)
            log.info(f"  ✓ Repaired: {product.product_name}")
        else:
            failed.append(product.product_name)
            errors.append(f"MSI repair failed for {product.product_name}: exit {result.returncode}")
            log.error(f"  ✗ Failed: {product.product_name} (exit {result.returncode})")

    return {"repaired": repaired, "failed": failed, "errors": errors}


def _repair_via_odt(
    products: list[DetectedOfficeProduct],
    *,
    strategy: RepairStrategy,
    dry_run: bool = False,
) -> dict[str, object]:
    """!
    @brief Repair products using ODT configuration.
    """
    log = logging_ext.get_human_logger()
    repaired: list[str] = []
    failed: list[str] = []
    errors: list[str] = []

    if not products:
        return {"repaired": repaired, "failed": failed, "errors": errors}

    # Determine preset based on strategy
    if strategy == RepairStrategy.FULL:
        preset_name = "full-repair"
    else:
        preset_name = "quick-repair"

    log.info(f"Running ODT repair with preset: {preset_name}")

    result = repair_module.run_oem_config(preset_name, dry_run=dry_run)

    if result.returncode == 0 or result.skipped:
        repaired.extend(p.product_name for p in products)
    else:
        failed.extend(p.product_name for p in products)
        errors.append(f"ODT repair failed: {result.stderr or result.error}")

    return {
        "repaired": repaired,
        "failed": failed,
        "errors": errors,
        "result": {
            "success": result.returncode == 0,
            "return_code": result.returncode,
        },
    }


# ---------------------------------------------------------------------------
# Convenience Functions
# ---------------------------------------------------------------------------


def quick_auto_repair(*, dry_run: bool = False) -> AutoRepairResult:
    """!
    @brief Run quick auto-repair on all detected Office installations.
    @param dry_run Simulate without executing.
    @returns AutoRepairResult with repair outcome.
    """
    return execute_auto_repair(
        strategy=RepairStrategy.QUICK,
        dry_run=dry_run,
    )


def full_auto_repair(*, dry_run: bool = False) -> AutoRepairResult:
    """!
    @brief Run full auto-repair on all detected Office installations.
    @param dry_run Simulate without executing.
    @returns AutoRepairResult with repair outcome.
    """
    return execute_auto_repair(
        strategy=RepairStrategy.FULL,
        dry_run=dry_run,
    )


def repair_c2r_quick(
    *, culture: str | None = None, dry_run: bool = False
) -> repair_module.RepairResult:
    """!
    @brief Quick repair C2R Office using OfficeClickToRun.exe.
    @param culture Language code (auto-detected if None).
    @param dry_run Simulate without executing.
    @returns RepairResult from repair module.
    """
    return repair_module.quick_repair(culture=culture, dry_run=dry_run)


def repair_c2r_full(
    *, culture: str | None = None, dry_run: bool = False
) -> repair_module.RepairResult:
    """!
    @brief Full online repair C2R Office using OfficeClickToRun.exe.
    @param culture Language code (auto-detected if None).
    @param dry_run Simulate without executing.
    @returns RepairResult from repair module.
    """
    return repair_module.full_repair(culture=culture, dry_run=dry_run)


def repair_via_odt_config(
    config_path: str | Path,
    *,
    dry_run: bool = False,
    timeout: int = 3600,
) -> CommandResult:
    """!
    @brief Repair/reconfigure Office using a custom ODT configuration.
    @param config_path Path to the XML configuration file.
    @param dry_run Simulate without executing.
    @param timeout Command timeout in seconds.
    @returns CommandResult with execution details.
    """
    return repair_module.reconfigure_office(
        Path(config_path),
        dry_run=dry_run,
        timeout=timeout,
    )


__all__ = [
    "RepairMethod",
    "RepairStrategy",
    "DetectedOfficeProduct",
    "AutoRepairPlan",
    "AutoRepairResult",
    "detect_office_products",
    "create_repair_plan",
    "execute_auto_repair",
    "quick_auto_repair",
    "full_auto_repair",
    "repair_c2r_quick",
    "repair_c2r_full",
    "repair_via_odt_config",
]
