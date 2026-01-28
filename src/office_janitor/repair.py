"""!
@brief Office Click-to-Run repair orchestration utilities.
@details Provides automated Quick Repair and Full Online Repair capabilities
for Office C2R installations using OfficeClicktoRun.exe. This module aligns
with the Office Deployment Tool documentation for programmatic repair scenarios.

ODT-related functionality (setup.exe, XML generation, OEM configs) has been
moved to :mod:`repair_odt` but is re-exported here for backwards compatibility.

@see https://docs.microsoft.com/en-us/deployoffice/overview-office-deployment-tool
@see https://itpro.outsidesys.com/2016/05/18/deploying-office-365-click-to-run/
"""

from __future__ import annotations

import re
from collections.abc import Mapping, MutableMapping
from dataclasses import dataclass
from enum import Enum
from pathlib import Path

from . import command_runner, constants, logging_ext, processes, registry_tools

# Import LogTailer for use in run_repair
from .repair_odt import LogTailer

# ---------------------------------------------------------------------------
# Constants and Enumerations
# ---------------------------------------------------------------------------


class RepairType(Enum):
    """!
    @brief Enumeration of supported Office repair strategies.
    @details The Quick Repair runs locally without network access and fixes
    common issues. The Full Repair reinstalls Office components from online
    sources and is more thorough but requires internet connectivity.
    """

    QUICK = "QuickRepair"
    FULL = "FullRepair"


class DisplayLevel(Enum):
    """!
    @brief UI visibility options for repair operations.
    """

    SILENT = "False"
    VISIBLE = "True"


# Default paths where OfficeClickToRun.exe is typically located
OFFICECLICKTORUN_CANDIDATES = (
    Path(r"C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"),
    Path(r"C:\Program Files (x86)\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"),
)
"""!
@brief Standard filesystem locations checked for OfficeClickToRun.exe.
"""

# Bundled executable in the oem folder (relative to package root)
BUNDLED_OFFICECLICKTORUN = "oem/OfficeClickToRun.exe"
"""!
@brief Path to bundled OfficeClickToRun.exe relative to the package root.
"""

# Platform identifiers
PLATFORM_X86 = "x86"
PLATFORM_X64 = "x64"

# Default culture/language codes
DEFAULT_CULTURE = "en-us"

# Supported language codes for Office repair
SUPPORTED_CULTURES = (
    "ar-sa",  # Arabic (Saudi Arabia)
    "bg-bg",  # Bulgarian
    "zh-cn",  # Chinese (Simplified)
    "zh-tw",  # Chinese (Traditional)
    "hr-hr",  # Croatian
    "cs-cz",  # Czech
    "da-dk",  # Danish
    "nl-nl",  # Dutch
    "en-us",  # English (US)
    "en-gb",  # English (UK)
    "et-ee",  # Estonian
    "fi-fi",  # Finnish
    "fr-fr",  # French
    "de-de",  # German
    "el-gr",  # Greek
    "he-il",  # Hebrew
    "hi-in",  # Hindi
    "hu-hu",  # Hungarian
    "id-id",  # Indonesian
    "it-it",  # Italian
    "ja-jp",  # Japanese
    "kk-kz",  # Kazakh
    "ko-kr",  # Korean
    "lv-lv",  # Latvian
    "lt-lt",  # Lithuanian
    "ms-my",  # Malay
    "nb-no",  # Norwegian (Bokmål)
    "pl-pl",  # Polish
    "pt-br",  # Portuguese (Brazil)
    "pt-pt",  # Portuguese (Portugal)
    "ro-ro",  # Romanian
    "ru-ru",  # Russian
    "sr-latn-rs",  # Serbian (Latin)
    "sk-sk",  # Slovak
    "sl-si",  # Slovenian
    "es-es",  # Spanish (Spain)
    "sv-se",  # Swedish
    "th-th",  # Thai
    "tr-tr",  # Turkish
    "uk-ua",  # Ukrainian
    "vi-vn",  # Vietnamese
)
"""!
@brief Language culture codes supported for Office repair operations.
"""

# Timeout values
REPAIR_TIMEOUT_QUICK = 1800  # 30 minutes
REPAIR_TIMEOUT_FULL = 7200  # 2 hours (online repair may take longer)

# Verification settings
REPAIR_VERIFICATION_ATTEMPTS = 3
REPAIR_VERIFICATION_DELAY = 5.0


# ---------------------------------------------------------------------------
# Data Classes
# ---------------------------------------------------------------------------


@dataclass
class RepairConfig:
    """!
    @brief Configuration parameters for an Office repair operation.
    @details Encapsulates all options required to invoke OfficeClickToRun.exe
    with the repair scenario. Use the factory methods for common configurations.
    """

    repair_type: RepairType = RepairType.QUICK
    platform: str = PLATFORM_X64
    culture: str = DEFAULT_CULTURE
    force_app_shutdown: bool = True
    display_level: DisplayLevel = DisplayLevel.SILENT
    timeout: int | None = None
    custom_exe_path: Path | None = None

    def __post_init__(self) -> None:
        """Validate culture code format."""
        if self.culture.lower() not in SUPPORTED_CULTURES:
            # Allow any ll-cc format even if not in our list
            if not re.match(r"^[a-z]{2,3}(-[a-z]{2,4})?$", self.culture.lower()):
                raise ValueError(f"Invalid culture code format: {self.culture}")

    @property
    def effective_timeout(self) -> int:
        """!
        @brief Return the timeout based on repair type if not explicitly set.
        """
        if self.timeout is not None:
            return self.timeout
        return REPAIR_TIMEOUT_FULL if self.repair_type == RepairType.FULL else REPAIR_TIMEOUT_QUICK

    def to_command_args(self) -> tuple[str, ...]:
        """!
        @brief Build the command-line arguments for OfficeClickToRun.exe.
        @returns Tuple of argument strings in the expected format.
        """
        return (
            "scenario=Repair",
            f"platform={self.platform}",
            f"culture={self.culture}",
            f"RepairType={self.repair_type.value}",
            f"forceappshutdown={str(self.force_app_shutdown)}",
            f"DisplayLevel={self.display_level.value}",
        )

    @classmethod
    def quick_repair(
        cls,
        *,
        platform: str | None = None,
        culture: str = DEFAULT_CULTURE,
        force_shutdown: bool = True,
        silent: bool = True,
    ) -> RepairConfig:
        """!
        @brief Factory method for creating a Quick Repair configuration.
        @param platform Architecture (x86/x64). Auto-detected if None.
        @param culture Language code (default: en-us).
        @param force_shutdown Whether to force close Office apps.
        @param silent Run without UI.
        @returns Configured RepairConfig for Quick Repair.
        """
        return cls(
            repair_type=RepairType.QUICK,
            platform=platform or _detect_office_platform(),
            culture=culture,
            force_app_shutdown=force_shutdown,
            display_level=DisplayLevel.SILENT if silent else DisplayLevel.VISIBLE,
        )

    @classmethod
    def full_repair(
        cls,
        *,
        platform: str | None = None,
        culture: str = DEFAULT_CULTURE,
        force_shutdown: bool = True,
        silent: bool = True,
    ) -> RepairConfig:
        """!
        @brief Factory method for creating a Full Online Repair configuration.
        @param platform Architecture (x86/x64). Auto-detected if None.
        @param culture Language code (default: en-us).
        @param force_shutdown Whether to force close Office apps.
        @param silent Run without UI.
        @returns Configured RepairConfig for Full Repair.
        """
        return cls(
            repair_type=RepairType.FULL,
            platform=platform or _detect_office_platform(),
            culture=culture,
            force_app_shutdown=force_shutdown,
            display_level=DisplayLevel.SILENT if silent else DisplayLevel.VISIBLE,
        )


@dataclass
class RepairResult:
    """!
    @brief Outcome of a repair operation.
    @details Contains information about whether the repair succeeded, any errors
    encountered, and duration metrics.
    """

    success: bool
    repair_type: RepairType
    return_code: int
    duration: float
    stdout: str = ""
    stderr: str = ""
    error_message: str | None = None
    skipped: bool = False
    timed_out: bool = False
    exe_path: str = ""

    @property
    def summary(self) -> str:
        """!
        @brief Human-readable summary of the repair outcome.
        """
        if self.skipped:
            return f"{self.repair_type.value} skipped (dry-run mode)"
        if self.timed_out:
            return f"{self.repair_type.value} timed out after {self.duration:.1f}s"
        if self.success:
            return f"{self.repair_type.value} completed successfully in {self.duration:.1f}s"
        error_msg = self.error_message or "Unknown error"
        return f"{self.repair_type.value} failed with code {self.return_code}: {error_msg}"


# ---------------------------------------------------------------------------
# Detection Utilities
# ---------------------------------------------------------------------------


def _detect_office_platform() -> str:
    """!
    @brief Detect the installed Office platform architecture.
    @details Reads the Office C2R registry configuration to determine whether
    x86 or x64 Office is installed. Falls back to x64 if unable to determine.
    @returns Platform string: 'x86' or 'x64'.
    """
    c2r_config_path = r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    try:
        platform_value = registry_tools.get_value(constants.HKLM, c2r_config_path, "Platform")
        if platform_value:
            if "x86" in str(platform_value).lower() or "32" in str(platform_value):
                return PLATFORM_X86
            return PLATFORM_X64
    except Exception:
        pass
    # Fallback: check Program Files paths
    if Path(r"C:\Program Files (x86)\Microsoft Office").exists():
        return PLATFORM_X86
    return PLATFORM_X64


def _detect_office_culture() -> str:
    """!
    @brief Detect the installed Office language/culture.
    @details Reads the Office C2R registry configuration to determine the
    primary language. Falls back to system locale or en-us.
    @returns Culture code string (e.g., 'en-us').
    """
    c2r_config_path = r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    try:
        # Try ClientCulture first
        culture = registry_tools.get_value(constants.HKLM, c2r_config_path, "ClientCulture")
        if culture:
            return str(culture).lower()
        # Try CDNBaseUrl which contains culture info
        cdn_url = registry_tools.get_value(constants.HKLM, c2r_config_path, "CDNBaseUrl")
        if cdn_url:
            # Extract culture from URL pattern
            match = re.search(r"/([a-z]{2}-[a-z]{2})/", str(cdn_url).lower())
            if match:
                return match.group(1)
    except Exception:
        pass

    # Fallback: try to get system locale
    try:
        import locale

        # Use getlocale for Python 3.11+ compatibility
        try:
            system_locale = locale.getlocale()[0]
        except Exception:
            system_locale = None
        if system_locale:
            # Convert xx_XX to xx-xx format
            culture = system_locale.lower().replace("_", "-")
            if re.match(r"^[a-z]{2}-[a-z]{2}$", culture):
                return culture
    except Exception:
        pass

    return DEFAULT_CULTURE


def find_officeclicktorun_exe(custom_path: Path | None = None) -> Path | None:
    """!
    @brief Locate the OfficeClickToRun.exe executable.
    @details Searches in order: custom path, standard system locations, then
    bundled executable in the oem folder.
    @param custom_path Optional explicit path to use.
    @returns Path to executable if found, None otherwise.
    """
    if custom_path and custom_path.exists():
        return custom_path

    # Check standard system locations
    for candidate in OFFICECLICKTORUN_CANDIDATES:
        if candidate.exists():
            return candidate

    # Check bundled executable
    try:
        import sys

        if getattr(sys, "frozen", False):
            # PyInstaller bundle
            base_path = Path(sys._MEIPASS)  # PyInstaller
        else:
            # Development - look relative to package
            base_path = Path(__file__).parent.parent.parent
        bundled = base_path / BUNDLED_OFFICECLICKTORUN
        if bundled.exists():
            return bundled
    except Exception:
        pass

    return None


def is_c2r_office_installed() -> bool:
    """!
    @brief Check if a Click-to-Run Office installation exists.
    @returns True if C2R Office is detected, False otherwise.
    """
    c2r_config_path = r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    try:
        platform = registry_tools.get_value(constants.HKLM, c2r_config_path, "Platform")
        return platform is not None
    except Exception:
        return False


def get_installed_c2r_info() -> Mapping[str, str | None]:
    """!
    @brief Retrieve information about the installed C2R Office.
    @returns Dictionary with version, platform, culture, and product info.
    """
    c2r_config_path = r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    info: MutableMapping[str, str | None] = {
        "version": None,
        "platform": None,
        "culture": None,
        "product_ids": None,
        "update_channel": None,
    }

    try:
        info["version"] = registry_tools.get_value(
            constants.HKLM, c2r_config_path, "VersionToReport"
        )
        info["platform"] = registry_tools.get_value(constants.HKLM, c2r_config_path, "Platform")
        info["culture"] = registry_tools.get_value(constants.HKLM, c2r_config_path, "ClientCulture")
        info["product_ids"] = registry_tools.get_value(
            constants.HKLM, c2r_config_path, "ProductReleaseIds"
        )
        info["update_channel"] = registry_tools.get_value(
            constants.HKLM, c2r_config_path, "CDNBaseUrl"
        )
    except Exception:
        pass

    return info


# ---------------------------------------------------------------------------
# Repair Execution
# ---------------------------------------------------------------------------


def run_repair(
    config: RepairConfig,
    *,
    dry_run: bool = False,
    close_office_apps: bool = True,
    log_callback: object | None = None,
) -> RepairResult:
    """!
    @brief Execute an Office repair operation.
    @details Locates the OfficeClickToRun.exe executable and runs it with the
    specified repair configuration. Optionally closes running Office applications
    before starting the repair.
    @param config RepairConfig instance with repair parameters.
    @param dry_run If True, log actions without executing.
    @param close_office_apps If True, terminate Office processes first.
    @param log_callback Optional callback function(str) to receive log output.
    @returns RepairResult with outcome information.
    """
    log = logging_ext.get_human_logger()
    mlog = logging_ext.get_machine_logger()

    # Find the executable
    exe_path = find_officeclicktorun_exe(config.custom_exe_path)
    if exe_path is None:
        error_msg = "OfficeClickToRun.exe not found"
        log.error(error_msg)
        mlog.info(
            "repair_error",
            extra={"event": "repair_error", "error": error_msg},
        )
        return RepairResult(
            success=False,
            repair_type=config.repair_type,
            return_code=-1,
            duration=0.0,
            error_message=error_msg,
        )

    # Close Office applications if requested
    if close_office_apps and config.force_app_shutdown:
        if not dry_run:
            _close_office_applications()

    # Build command
    command = [str(exe_path)] + list(config.to_command_args())

    log.info(f"Starting {config.repair_type.value} repair...")
    mlog.info(
        "repair_start",
        extra={
            "event": "repair_start",
            "repair_type": config.repair_type.value,
            "platform": config.platform,
            "culture": config.culture,
            "exe_path": str(exe_path),
            "command": command,
            "dry_run": dry_run,
        },
    )

    if dry_run:
        log.info(f"[DRY-RUN] Would execute: {' '.join(command)}")
        return RepairResult(
            success=True,
            repair_type=config.repair_type,
            return_code=0,
            duration=0.0,
            skipped=True,
            exe_path=str(exe_path),
        )

    # Execute repair with log tailing
    log.info("Tailing Office repair logs from %temp%...")
    with LogTailer(output_callback=log_callback):
        result = command_runner.run_command(
            command,
            event="repair_exec",
            timeout=config.effective_timeout,
            dry_run=dry_run,
        )

    repair_result = RepairResult(
        success=result.returncode == 0,
        repair_type=config.repair_type,
        return_code=result.returncode,
        duration=result.duration,
        stdout=result.stdout,
        stderr=result.stderr,
        error_message=result.error if result.error else (result.stderr if result.stderr else None),
        timed_out=result.timed_out,
        exe_path=str(exe_path),
    )

    mlog.info(
        "repair_complete",
        extra={
            "event": "repair_complete",
            "success": repair_result.success,
            "repair_type": config.repair_type.value,
            "return_code": result.returncode,
            "duration": result.duration,
            "timed_out": result.timed_out,
        },
    )

    if repair_result.success:
        log.info(f"{config.repair_type.value} completed successfully")
    else:
        log.error(f"{config.repair_type.value} failed: {repair_result.error_message}")

    return repair_result


def quick_repair(
    *,
    culture: str | None = None,
    silent: bool = True,
    dry_run: bool = False,
    log_callback: object | None = None,
) -> RepairResult:
    """!
    @brief Convenience function to run a Quick Repair.
    @details Detects platform and culture automatically if not specified.
    Quick Repair runs locally and fixes common issues without internet.
    @param culture Language code (auto-detected if None).
    @param silent Run without UI.
    @param dry_run Simulate without executing.
    @param log_callback Optional callback function(str) to receive log output.
    @returns RepairResult with outcome information.
    """
    config = RepairConfig.quick_repair(
        culture=culture or _detect_office_culture(),
        silent=silent,
    )
    return run_repair(config, dry_run=dry_run, log_callback=log_callback)


def full_repair(
    *,
    culture: str | None = None,
    silent: bool = True,
    dry_run: bool = False,
    log_callback: object | None = None,
) -> RepairResult:
    """!
    @brief Convenience function to run a Full Online Repair.
    @details Detects platform and culture automatically if not specified.
    Full Repair reinstalls Office from online sources and requires internet.
    WARNING: Full repair may reinstall previously excluded applications.
    @param culture Language code (auto-detected if None).
    @param silent Run without UI.
    @param dry_run Simulate without executing.
    @param log_callback Optional callback function(str) to receive log output.
    @returns RepairResult with outcome information.
    """
    config = RepairConfig.full_repair(
        culture=culture or _detect_office_culture(),
        silent=silent,
    )
    return run_repair(config, dry_run=dry_run, log_callback=log_callback)


def _close_office_applications() -> None:
    """!
    @brief Terminate running Office applications.
    @details Gracefully closes common Office processes before repair.
    """
    log = logging_ext.get_human_logger()
    office_processes = list(constants.DEFAULT_OFFICE_PROCESSES)
    try:
        processes.terminate_office_processes(office_processes)
    except Exception as e:
        log.debug(f"Could not terminate Office processes: {e}")


# ---------------------------------------------------------------------------
# Re-exports from repair_odt for backwards compatibility
# ---------------------------------------------------------------------------

from .repair_odt import (  # noqa: E402
    BUNDLED_ODT_SETUP,
    BUNDLED_OEM_DIR,
    ODT_LOG_PATTERNS,
    ODT_SETUP_CANDIDATES,
    OEM_CONFIG_PRESETS,
    find_odt_setup_exe,
    generate_repair_config_xml,
    get_oem_config_path,
    list_oem_configs,
    reconfigure_office,
    run_oem_config,
)

__all__ = [
    # Core repair types and config
    "RepairType",
    "DisplayLevel",
    "RepairConfig",
    "RepairResult",
    # Constants
    "OFFICECLICKTORUN_CANDIDATES",
    "BUNDLED_OFFICECLICKTORUN",
    "PLATFORM_X86",
    "PLATFORM_X64",
    "DEFAULT_CULTURE",
    "SUPPORTED_CULTURES",
    "REPAIR_TIMEOUT_QUICK",
    "REPAIR_TIMEOUT_FULL",
    # Detection
    "find_officeclicktorun_exe",
    "is_c2r_office_installed",
    "get_installed_c2r_info",
    # Repair execution
    "run_repair",
    "quick_repair",
    "full_repair",
    # Re-exports from repair_odt
    "LogTailer",
    "ODT_LOG_PATTERNS",
    "ODT_SETUP_CANDIDATES",
    "BUNDLED_ODT_SETUP",
    "BUNDLED_OEM_DIR",
    "OEM_CONFIG_PRESETS",
    "find_odt_setup_exe",
    "reconfigure_office",
    "generate_repair_config_xml",
    "get_oem_config_path",
    "list_oem_configs",
    "run_oem_config",
]
