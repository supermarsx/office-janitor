"""!
@brief Office Deployment Tool (ODT) integration for C2R management.
@details Provides utilities for downloading ODT, building configuration XML,
and executing ODT-based Office removal operations. VBS equivalents are found
in OffScrubC2R.vbs functions like BuildRemoveXml, DownloadODT, and
UninstallOfficeC2R.
"""

from __future__ import annotations

import tempfile
import urllib.error
import urllib.request
from pathlib import Path

from . import command_runner, logging_ext

# ---------------------------------------------------------------------------
# ODT Download URLs and Templates
# ---------------------------------------------------------------------------

ODT_DOWNLOAD_URLS = {
    16: "https://officecdn.microsoft.com/pr/wsus/setup.exe",
    15: (
        "https://download.microsoft.com/download/2/7/A/"
        "27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_x86_5031-1000.exe"
    ),
}
"""!
@brief URLs for downloading ODT setup.exe by Office version.
@details Version 16 covers Office 365/2016/2019/2021/2024.
Version 15 covers Office 2013.
"""

ODT_REMOVE_XML_TEMPLATE = """<Configuration>
  <Remove All="TRUE" />
  <Display Level="{level}" AcceptEULA="TRUE" />
</Configuration>
"""
"""!
@brief Template for ODT configuration XML to remove all Office products.
"""

ODT_REMOVE_PRODUCT_XML_TEMPLATE = """<Configuration>
  <Remove>
    <Product ID="{product_id}" />
  </Remove>
  <Display Level="{level}" AcceptEULA="TRUE" />
</Configuration>
"""
"""!
@brief Template for ODT configuration XML to remove a specific product.
"""

ODT_TIMEOUT = 3600
"""!
@brief Default timeout for ODT operations in seconds.
"""

# ---------------------------------------------------------------------------
# Standard ODT/Setup.exe Locations
# ---------------------------------------------------------------------------

C2R_SETUP_CANDIDATES = (
    Path(r"C:\Program Files\Common Files\Microsoft Shared\ClickToRun\setup.exe"),
    Path(r"C:\Program Files (x86)\Common Files\Microsoft Shared\ClickToRun\setup.exe"),
)
"""!
@brief Default filesystem locations checked for ``setup.exe``.
"""


# ---------------------------------------------------------------------------
# XML Building Functions
# ---------------------------------------------------------------------------


def build_remove_xml(
    output_path: Path | str,
    *,
    quiet: bool = True,
) -> Path:
    """!
    @brief Build ODT configuration XML for complete Office removal.
    @details VBS equivalent: BuildRemoveXml in OffScrubC2R.vbs.
    @param output_path Path where the XML file should be written.
    @param quiet If True, use silent display level; otherwise use Full.
    @returns Path to the written XML file.
    """
    level = "None" if quiet else "Full"
    content = ODT_REMOVE_XML_TEMPLATE.format(level=level)

    path = Path(output_path)
    path.write_text(content, encoding="utf-8")

    return path


def build_remove_product_xml(
    output_path: Path | str,
    product_id: str,
    *,
    quiet: bool = True,
) -> Path:
    """!
    @brief Build ODT configuration XML for removing a specific product.
    @param output_path Path where the XML file should be written.
    @param product_id The Office product ID to remove (e.g., "O365ProPlusRetail").
    @param quiet If True, use silent display level; otherwise use Full.
    @returns Path to the written XML file.
    """
    level = "None" if quiet else "Full"
    content = ODT_REMOVE_PRODUCT_XML_TEMPLATE.format(product_id=product_id, level=level)

    path = Path(output_path)
    path.write_text(content, encoding="utf-8")

    return path


def build_custom_remove_xml(
    output_path: Path | str,
    product_ids: list[str],
    *,
    quiet: bool = True,
    force_app_shutdown: bool = False,
) -> Path:
    """!
    @brief Build ODT configuration XML for removing multiple products.
    @param output_path Path where the XML file should be written.
    @param product_ids List of Office product IDs to remove.
    @param quiet If True, use silent display level; otherwise use Full.
    @param force_app_shutdown If True, add ForceAppShutdown property.
    @returns Path to the written XML file.
    """
    level = "None" if quiet else "Full"

    products_xml = "\n".join(f'    <Product ID="{pid}" />' for pid in product_ids)

    force_shutdown = ""
    if force_app_shutdown:
        force_shutdown = '\n  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />'

    content = f"""<Configuration>
  <Remove>
{products_xml}
  </Remove>
  <Display Level="{level}" AcceptEULA="TRUE" />{force_shutdown}
</Configuration>
"""

    path = Path(output_path)
    path.write_text(content, encoding="utf-8")

    return path


# ---------------------------------------------------------------------------
# ODT Download Functions
# ---------------------------------------------------------------------------


def download_odt(
    version: int = 16,
    dest_dir: Path | str | None = None,
    *,
    dry_run: bool = False,
) -> Path | None:
    """!
    @brief Download Office Deployment Tool if not available locally.
    @details VBS equivalent: DownloadODT in OffScrubC2R.vbs.
    Uses urllib to fetch ODT setup.exe from Microsoft CDN.
    @param version Office version (15 or 16). Version 16 covers 365/2016/2019/2021/2024.
    @param dest_dir Directory to save the downloaded file. Defaults to temp dir.
    @param dry_run If True, only log what would be done without downloading.
    @returns Path to downloaded setup.exe, or None if download failed.
    """
    human_logger = logging_ext.get_human_logger()

    url = ODT_DOWNLOAD_URLS.get(version)
    if not url:
        human_logger.error("No ODT download URL for Office version %d", version)
        return None

    if dest_dir is None:
        dest_dir = Path(tempfile.gettempdir())
    else:
        dest_dir = Path(dest_dir)

    dest_path = dest_dir / "setup.exe"

    if dry_run:
        human_logger.info("[DRY-RUN] Would download ODT from: %s to %s", url, dest_path)
        return dest_path

    human_logger.info("Downloading ODT from: %s", url)

    try:
        # Add User-Agent to avoid potential blocks
        request = urllib.request.Request(
            url,
            headers={"User-Agent": "OfficeJanitor/1.0"},
        )
        with urllib.request.urlopen(request, timeout=60) as response:
            content = response.read()

        dest_path.write_bytes(content)
        human_logger.info("Downloaded ODT to: %s (%d bytes)", dest_path, len(content))
        return dest_path

    except urllib.error.URLError as exc:
        human_logger.error("Failed to download ODT: %s", exc)
        return None
    except OSError as exc:
        human_logger.error("Failed to save ODT: %s", exc)
        return None


def find_local_odt() -> Path | None:
    """!
    @brief Find local ODT setup.exe in standard locations.
    @returns Path to setup.exe or None if not found.
    """
    human_logger = logging_ext.get_human_logger()

    for candidate in C2R_SETUP_CANDIDATES:
        if candidate.exists():
            human_logger.debug("Found local ODT at: %s", candidate)
            return candidate

    return None


def find_or_download_odt(
    version: int = 16,
    *,
    dry_run: bool = False,
) -> Path | None:
    """!
    @brief Find local ODT setup.exe or download if not found.
    @param version Office version for download URL selection.
    @param dry_run If True, don't actually download.
    @returns Path to setup.exe or None if unavailable.
    """
    human_logger = logging_ext.get_human_logger()

    # First check standard locations
    local_odt = find_local_odt()
    if local_odt:
        return local_odt

    # Not found locally - download
    human_logger.info("ODT not found locally, attempting download...")
    return download_odt(version, dry_run=dry_run)


# ---------------------------------------------------------------------------
# ODT Execution Functions
# ---------------------------------------------------------------------------


def uninstall_via_odt(
    odt_path: Path | str,
    config_xml: Path | str | None = None,
    *,
    quiet: bool = True,
    dry_run: bool = False,
    timeout: int = ODT_TIMEOUT,
) -> int:
    """!
    @brief Execute ODT-based Office uninstall.
    @details VBS equivalent: UninstallOfficeC2R using ODT in OffScrubC2R.vbs.
    @param odt_path Path to ODT setup.exe.
    @param config_xml Path to removal XML. If None, creates a temporary one.
    @param quiet If True, use silent mode.
    @param dry_run If True, only log what would be done.
    @param timeout Command timeout in seconds.
    @returns Exit code from setup.exe (0 = success).
    """
    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    odt_path = Path(odt_path)
    if not odt_path.exists():
        human_logger.error("ODT setup.exe not found: %s", odt_path)
        return 1

    # Build config XML if not provided
    if config_xml is None:
        temp_dir = Path(tempfile.gettempdir())
        config_xml = build_remove_xml(temp_dir / "RemoveAll.xml", quiet=quiet)
    else:
        config_xml = Path(config_xml)

    command = [str(odt_path), "/configure", str(config_xml)]

    if dry_run:
        human_logger.info("[DRY-RUN] Would execute: %s", " ".join(command))
        return 0

    human_logger.info("Executing ODT removal: %s", " ".join(command))
    machine_logger.info(
        "odt_uninstall_start",
        extra={
            "event": "odt_uninstall_start",
            "odt_path": str(odt_path),
            "config_xml": str(config_xml),
        },
    )

    result = command_runner.run_command(
        command,
        timeout=timeout,
        event="odt_uninstall",
    )

    machine_logger.info(
        "odt_uninstall_complete",
        extra={
            "event": "odt_uninstall_complete",
            "exit_code": result.returncode,
        },
    )

    if result.returncode == 0:
        human_logger.info("ODT removal completed successfully")
    else:
        human_logger.warning("ODT removal exited with code: %d", result.returncode)

    return result.returncode


def uninstall_product_via_odt(
    product_id: str,
    odt_path: Path | str | None = None,
    *,
    quiet: bool = True,
    dry_run: bool = False,
    timeout: int = ODT_TIMEOUT,
) -> int:
    """!
    @brief Uninstall a specific Office product via ODT.
    @param product_id The Office product ID to remove (e.g., "O365ProPlusRetail").
    @param odt_path Path to ODT setup.exe. If None, will try to find or download.
    @param quiet If True, use silent mode.
    @param dry_run If True, only log what would be done.
    @param timeout Command timeout in seconds.
    @returns Exit code from setup.exe (0 = success, 1 = error).
    """
    human_logger = logging_ext.get_human_logger()

    # Find ODT if not provided
    if odt_path is None:
        odt_path = find_or_download_odt(dry_run=dry_run)
        if odt_path is None:
            human_logger.error("Cannot find or download ODT setup.exe")
            return 1
    else:
        odt_path = Path(odt_path)

    # Build product-specific removal XML
    temp_dir = Path(tempfile.gettempdir())
    config_xml = build_remove_product_xml(
        temp_dir / f"Remove_{product_id}.xml",
        product_id,
        quiet=quiet,
    )

    return uninstall_via_odt(
        odt_path,
        config_xml,
        quiet=quiet,
        dry_run=dry_run,
        timeout=timeout,
    )


def uninstall_all_via_odt(
    *,
    quiet: bool = True,
    dry_run: bool = False,
    timeout: int = ODT_TIMEOUT,
) -> int:
    """!
    @brief Uninstall all Office products via ODT.
    @param quiet If True, use silent mode.
    @param dry_run If True, only log what would be done.
    @param timeout Command timeout in seconds.
    @returns Exit code from setup.exe (0 = success, 1 = error).
    """
    human_logger = logging_ext.get_human_logger()

    odt_path = find_or_download_odt(dry_run=dry_run)
    if odt_path is None:
        human_logger.error("Cannot find or download ODT setup.exe")
        return 1

    return uninstall_via_odt(
        odt_path,
        config_xml=None,  # Will build RemoveAll.xml
        quiet=quiet,
        dry_run=dry_run,
        timeout=timeout,
    )
