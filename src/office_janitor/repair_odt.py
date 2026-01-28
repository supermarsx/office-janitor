"""!
@brief Office Deployment Tool (ODT) repair and reconfiguration utilities.
@details Provides ODT setup.exe integration for reconfiguring Office installations,
log tailing for setup operations, and OEM configuration preset management.

@see https://docs.microsoft.com/en-us/deployoffice/overview-office-deployment-tool
"""

from __future__ import annotations

import glob
import os
import threading
import time
from collections.abc import Sequence
from pathlib import Path

from . import command_runner, constants, logging_ext, registry_tools
from .exec_utils import CommandResult

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

ODT_SETUP_CANDIDATES = (
    Path(r"C:\Program Files\Common Files\Microsoft Shared\ClickToRun\setup.exe"),
    Path(r"C:\Program Files (x86)\Common Files\Microsoft Shared\ClickToRun\setup.exe"),
)
"""!
@brief Standard paths for ODT setup.exe used for reconfigure operations.
"""

BUNDLED_ODT_SETUP = "oem/setup.exe"
"""!
@brief Path to bundled setup.exe relative to the package root.
"""

# Log file patterns for ODT/Office setup
ODT_LOG_PATTERNS = [
    "%temp%/SetupOffice*.log",
    "%temp%/OfficeClickToRun*.log",
    "%temp%/Office*.log",
]
"""!
@brief Glob patterns for Office setup log files in %temp%.
"""

# OEM configuration presets mapping (name -> filename)
OEM_CONFIG_PRESETS: dict[str, str] = {
    "full-removal": "FullRemoval.xml",
    "quick-repair": "QuickRepair.xml",
    "full-repair": "FullRepair.xml",
    "proplus-x64": "Repair_ProPlus_x64.xml",
    "proplus-x86": "Repair_ProPlus_x86.xml",
    "proplus-visio-project": "Repair_ProPlus_Visio_Project.xml",
    "business-x64": "Repair_Business_x64.xml",
    "office2019-x64": "Repair_Office2019_x64.xml",
    "office2021-x64": "Repair_Office2021_x64.xml",
    "office2024-x64": "Repair_Office2024_x64.xml",
    "multilang": "Repair_Multilang.xml",
    "shared-computer": "Repair_SharedComputer.xml",
    "interactive": "Repair_Interactive.xml",
}
"""!
@brief Mapping of OEM config preset names to their XML filenames in oem/ folder.
"""

BUNDLED_OEM_DIR = "oem"
"""!
@brief Relative path to the bundled OEM configurations directory.
"""

# Platform identifiers (imported from repair module constants)
PLATFORM_X86 = "x86"
PLATFORM_X64 = "x64"

# Default culture
DEFAULT_CULTURE = "en-us"


# ---------------------------------------------------------------------------
# OEM Config Path Resolution
# ---------------------------------------------------------------------------


def get_oem_config_path(preset_name: str) -> Path | None:
    """!
    @brief Resolve an OEM config preset name to its full file path.
    @param preset_name The preset name (key from OEM_CONFIG_PRESETS) or a direct filename.
    @returns Path to the XML config file if found, None otherwise.
    """
    import sys

    # Check if it's a preset name
    filename = OEM_CONFIG_PRESETS.get(preset_name, preset_name)

    # Determine base path for bundled resources
    try:
        if getattr(sys, "frozen", False):
            base_path = Path(sys._MEIPASS)  # PyInstaller
        else:
            base_path = Path(__file__).parent.parent.parent
        oem_path = base_path / BUNDLED_OEM_DIR / filename
        if oem_path.exists():
            return oem_path
    except Exception:
        pass

    # Also check if filename is an absolute path
    direct_path = Path(filename)
    if direct_path.is_absolute() and direct_path.exists():
        return direct_path

    return None


def list_oem_configs() -> list[tuple[str, str, bool]]:
    """!
    @brief List all available OEM configuration presets.
    @returns List of tuples: (preset_name, filename, exists).
    """
    import sys

    try:
        if getattr(sys, "frozen", False):
            base_path = Path(sys._MEIPASS)  # PyInstaller
        else:
            base_path = Path(__file__).parent.parent.parent
        oem_dir = base_path / BUNDLED_OEM_DIR
    except Exception:
        oem_dir = Path(BUNDLED_OEM_DIR)

    result: list[tuple[str, str, bool]] = []
    for preset_name, filename in OEM_CONFIG_PRESETS.items():
        exists = (oem_dir / filename).exists() if oem_dir.exists() else False
        result.append((preset_name, filename, exists))
    return result


# ---------------------------------------------------------------------------
# Log Tailer
# ---------------------------------------------------------------------------


class LogTailer:
    """!
    @brief Background log file tailer for Office setup operations.
    @details Monitors log files in %temp% and streams new content to console or a callback.
    """

    def __init__(
        self,
        patterns: list[str] | None = None,
        output_callback: object | None = None,
    ):
        self._patterns = patterns or ODT_LOG_PATTERNS
        self._output_callback = output_callback
        self._stop_event = threading.Event()
        self._thread: threading.Thread | None = None
        self._seen_files: set[str] = set()
        self._file_positions: dict[str, int] = {}
        self._log = logging_ext.get_human_logger()

    def _expand_pattern(self, pattern: str) -> str:
        """Expand environment variables in pattern."""
        return os.path.expandvars(pattern)

    def _find_log_files(self) -> list[Path]:
        """Find all matching log files."""
        files: list[Path] = []
        for pattern in self._patterns:
            expanded = self._expand_pattern(pattern)
            for match in glob.glob(expanded):
                files.append(Path(match))
        return files

    def _tail_file(self, filepath: Path) -> None:
        """Read and print new content from a log file."""
        try:
            str_path = str(filepath)
            if str_path not in self._file_positions:
                # New file - start from beginning if recently created, else end
                stat = filepath.stat()
                age = time.time() - stat.st_mtime
                if age < 5:  # File created in last 5 seconds
                    self._file_positions[str_path] = 0
                else:
                    self._file_positions[str_path] = stat.st_size

            pos = self._file_positions[str_path]
            current_size = filepath.stat().st_size

            if current_size > pos:
                with open(filepath, encoding="utf-8", errors="replace") as f:
                    f.seek(pos)
                    new_content = f.read()
                    if new_content.strip():
                        # Send output via callback or print
                        for line in new_content.splitlines():
                            if line.strip():
                                output_line = f"  [ODT] {line}"
                                if self._output_callback is not None:
                                    try:
                                        self._output_callback(output_line)
                                    except Exception:
                                        pass  # Ignore callback errors
                                else:
                                    print(output_line)
                    self._file_positions[str_path] = f.tell()

        except OSError:
            pass  # File may be locked or deleted

    def _monitor_loop(self) -> None:
        """Main monitoring loop."""
        while not self._stop_event.is_set():
            for filepath in self._find_log_files():
                self._tail_file(filepath)
            self._stop_event.wait(0.5)  # Check every 500ms

    def start(self) -> None:
        """Start the log tailer in a background thread."""
        if self._thread is not None:
            return
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self._thread.start()
        self._log.debug("Log tailer started")

    def stop(self) -> None:
        """Stop the log tailer."""
        if self._thread is None:
            return
        self._stop_event.set()
        self._thread.join(timeout=2.0)
        self._thread = None
        self._log.debug("Log tailer stopped")

    def __enter__(self) -> LogTailer:
        self.start()
        return self

    def __exit__(self, *args: object) -> None:
        self.stop()


# ---------------------------------------------------------------------------
# ODT Setup Utilities
# ---------------------------------------------------------------------------


def find_odt_setup_exe(custom_path: Path | None = None) -> Path | None:
    """!
    @brief Locate the Office Deployment Tool setup.exe.
    @param custom_path Optional explicit path to use.
    @returns Path to executable if found, None otherwise.
    """
    if custom_path and custom_path.exists():
        return custom_path

    for candidate in ODT_SETUP_CANDIDATES:
        if candidate.exists():
            return candidate

    # Check bundled executable
    try:
        import sys

        if getattr(sys, "frozen", False):
            base_path = Path(sys._MEIPASS)  # PyInstaller
        else:
            base_path = Path(__file__).parent.parent.parent
        bundled = base_path / BUNDLED_ODT_SETUP
        if bundled.exists():
            return bundled
    except Exception:
        pass

    return None


def reconfigure_office(
    config_xml_path: Path,
    *,
    dry_run: bool = False,
    timeout: int = 3600,
    log_callback: object | None = None,
    progress_callback: object | None = None,
) -> CommandResult:
    """!
    @brief Reconfigure Office installation using ODT setup.exe.
    @details Uses the /configure switch with an XML configuration file to
    modify the Office installation (add/remove apps, languages, etc.).
    @param config_xml_path Path to the configuration XML file.
    @param dry_run Simulate without executing.
    @param timeout Command timeout in seconds.
    @param log_callback Optional callback function(str) to receive log output.
    @param progress_callback Optional callback function(str) to receive progress updates.
    @returns CommandResult with execution details.
    """
    log = logging_ext.get_human_logger()
    mlog = logging_ext.get_machine_logger()

    setup_exe = find_odt_setup_exe()
    if setup_exe is None:
        log.error("ODT setup.exe not found")
        return CommandResult(
            command=[],
            returncode=-1,
            stdout="",
            stderr="ODT setup.exe not found",
            duration=0.0,
            error="ODT setup.exe not found",
        )

    if not config_xml_path.exists():
        log.error(f"Configuration XML not found: {config_xml_path}")
        return CommandResult(
            command=[],
            returncode=-1,
            stdout="",
            stderr=f"Configuration XML not found: {config_xml_path}",
            duration=0.0,
            error=f"Configuration XML not found: {config_xml_path}",
        )

    command = [str(setup_exe), "/configure", str(config_xml_path)]

    mlog.info(
        "reconfigure_start",
        extra={
            "event": "reconfigure_start",
            "config_xml": str(config_xml_path),
            "command": command,
        },
    )

    # Use log tailer to stream ODT logs to console
    if not dry_run:
        log.info("Tailing ODT logs from %temp%...")
        # Combine callbacks: send output to both log_callback and progress_callback
        combined_callback = None
        if log_callback or progress_callback:
            def combined_callback(line: str) -> None:
                if log_callback:
                    log_callback(line)
                if progress_callback:
                    progress_callback(line)
        
        with LogTailer(output_callback=combined_callback):
            result = command_runner.run_command(
                command,
                event="reconfigure_exec",
                timeout=timeout,
                dry_run=dry_run,
            )
        return result
    else:
        return command_runner.run_command(
            command,
            event="reconfigure_exec",
            timeout=timeout,
            dry_run=dry_run,
        )


# ---------------------------------------------------------------------------
# XML Configuration Generators
# ---------------------------------------------------------------------------


def _detect_office_platform() -> str:
    """!
    @brief Detect the installed Office platform architecture.
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


def _get_installed_c2r_info() -> dict[str, str | None]:
    """!
    @brief Retrieve information about the installed C2R Office.
    @returns Dictionary with version, platform, culture, and product info.
    """
    c2r_config_path = r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    info: dict[str, str | None] = {
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


def generate_repair_config_xml(
    output_path: Path,
    *,
    product_ids: Sequence[str] | None = None,
    language: str = DEFAULT_CULTURE,
    force_app_shutdown: bool = True,
) -> Path:
    """!
    @brief Generate an XML configuration for Office repair/reconfiguration.
    @details Creates a Configuration.xml suitable for use with ODT setup.exe
    in /configure mode to trigger a repair-like reconfiguration.
    @param output_path Where to write the XML file.
    @param product_ids List of product IDs (e.g., ['O365ProPlusRetail']).
    @param language Primary language code.
    @param force_app_shutdown Include FORCEAPPSHUTDOWN property.
    @returns Path to the generated XML file.
    """
    if product_ids is None:
        # Try to detect installed products
        info = _get_installed_c2r_info()
        if info.get("product_ids"):
            product_ids = [p.strip() for p in str(info["product_ids"]).split(",")]
        else:
            product_ids = ["O365ProPlusRetail"]

    products_xml = ""
    for pid in product_ids:
        products_xml += f"""    <Product ID="{pid}">
      <Language ID="{language}" />
    </Product>
"""

    # Determine edition based on platform
    platform = _detect_office_platform()
    edition = "64" if platform == PLATFORM_X64 else "32"

    xml_content = f"""<Configuration>
  <Add OfficeClientEdition="{edition}">
{products_xml}  </Add>
  <Updates Enabled="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Logging Level="Standard" Path="%temp%" />
  <Property Name="FORCEAPPSHUTDOWN" Value="{str(force_app_shutdown).upper()}" />
</Configuration>
"""

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(xml_content, encoding="utf-8")
    return output_path


def run_oem_config(
    preset_or_path: str,
    *,
    dry_run: bool = False,
    timeout: int = 3600,
    log_callback: object | None = None,
    progress_callback: object | None = None,
) -> CommandResult:
    """!
    @brief Execute an OEM configuration preset or custom XML config.
    @details Resolves preset names to bundled XML files and runs them via
    ODT setup.exe /configure. Can also accept direct paths to XML files.
    @param preset_or_path Preset name (from OEM_CONFIG_PRESETS) or path to XML file.
    @param dry_run Simulate without executing.
    @param timeout Command timeout in seconds.
    @param log_callback Optional callback function(str) to receive log output.
    @param progress_callback Optional callback function(str) to receive progress updates.
    @returns CommandResult with execution details.
    """
    log = logging_ext.get_human_logger()
    mlog = logging_ext.get_machine_logger()

    # Resolve the config path
    config_path = get_oem_config_path(preset_or_path)

    if config_path is None:
        error_msg = f"OEM config not found: {preset_or_path}"
        log.error(error_msg)
        available = [name for name, _, exists in list_oem_configs() if exists]
        if available:
            log.info(f"Available presets: {', '.join(available)}")
        return CommandResult(
            command=[],
            returncode=-1,
            stdout="",
            stderr=error_msg,
            duration=0.0,
            error=error_msg,
        )

    mlog.info(
        "oem_config_start",
        extra={
            "event": "oem_config_start",
            "preset": preset_or_path,
            "resolved_path": str(config_path),
        },
    )

    log.info(f"Running OEM config: {preset_or_path} -> {config_path}")

    return reconfigure_office(
        config_path,
        dry_run=dry_run,
        timeout=timeout,
        log_callback=log_callback,
        progress_callback=progress_callback,
    )


__all__ = [
    "BUNDLED_ODT_SETUP",
    "BUNDLED_OEM_DIR",
    "DEFAULT_CULTURE",
    "LogTailer",
    "ODT_LOG_PATTERNS",
    "ODT_SETUP_CANDIDATES",
    "OEM_CONFIG_PRESETS",
    "PLATFORM_X64",
    "PLATFORM_X86",
    "find_odt_setup_exe",
    "generate_repair_config_xml",
    "get_oem_config_path",
    "list_oem_configs",
    "reconfigure_office",
    "run_oem_config",
]
