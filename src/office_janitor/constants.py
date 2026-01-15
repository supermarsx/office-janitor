"""!
@brief Static data and enumerations for Office Janitor.
@details Centralises product identifiers, registry roots, Click-to-Run channel
metadata, uninstall command templates, and supported targets so detection and
scrub modules operate from a consistent catalogue of Microsoft Office
artifacts.
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from enum import IntFlag
from typing import TYPE_CHECKING, cast

if TYPE_CHECKING:  # pragma: no cover - type checking only
    import winreg as _winreg
else:
    try:  # pragma: no cover - Windows registry handles are optional on test hosts.
        import winreg as _winreg  # type: ignore[import-not-found]
    except ImportError:  # pragma: no cover - test scaffolding supplies substitutes.

        class _WinRegStub:
            HKEY_LOCAL_MACHINE = 0x80000002
            HKEY_CURRENT_USER = 0x80000001
            HKEY_CLASSES_ROOT = 0x80000000
            HKEY_USERS = 0x80000003

        _winreg = _WinRegStub()  # type: ignore[assignment]

winreg = _winreg  # re-export for callers

HKLM = cast(int, getattr(winreg, "HKEY_LOCAL_MACHINE", 0x80000002))
HKCU = cast(int, getattr(winreg, "HKEY_CURRENT_USER", 0x80000001))
HKCR = cast(int, getattr(winreg, "HKEY_CLASSES_ROOT", 0x80000000))
HKU = cast(int, getattr(winreg, "HKEY_USERS", 0x80000003))


REGISTRY_ROOTS: dict[str, int] = {
    "HKLM": HKLM,
    "HKCU": HKCU,
    "HKCR": HKCR,
    "HKU": HKU,
}


# ---------------------------------------------------------------------------
# Error Bitmask System (VBS-compatible return codes)
# ---------------------------------------------------------------------------


class ScrubErrorCode(IntFlag):
    """!
    @brief Bitmask error codes for scrub operations.
    @details VBS-compatible return codes from OffScrub*.vbs scripts.
        Multiple flags can be combined to indicate multiple error conditions.
    """

    SUCCESS = 0
    """No errors."""
    FAIL = 1
    """General failure."""
    REBOOT_REQUIRED = 2
    """Reboot needed to complete removal."""
    USER_CANCEL = 4
    """User cancelled the operation."""
    MSIEXEC_FAILED = 8
    """MSI uninstall stage failed."""
    CLEANUP_FAILED = 16
    """Post-uninstall cleanup failed."""
    INCOMPLETE = 32
    """Pending file renames require reboot."""
    RETRY_FAILED = 64
    """Second attempt (DCAF) also failed."""
    ELEVATION_DECLINED = 128
    """User declined elevation prompt."""
    ELEVATION_ERROR = 256
    """UAC elevation failed."""
    INIT_ERROR = 512
    """Script initialization failed."""
    RELAUNCH_ERROR = 1024
    """Elevated relaunch failed."""
    UNKNOWN = 2048
    """Unknown/unexpected error."""


SUPPORTED_VERSIONS = (
    "2003",
    "2007",
    "2010",
    "2013",
    "2016",
    "2019",
    "2021",
    "2024",
    "365",
)
"""!
@brief Ordered tuple of Office generations targeted by the scrubber.
"""

SUPPORTED_TARGETS = SUPPORTED_VERSIONS
"""!
@brief Alias for supported targets used by planning logic.
"""

SUPPORTED_COMPONENTS = (
    "visio",
    "project",
    "onenote",
)
"""!
@brief Optional components that callers may include explicitly.
"""

DEFAULT_OFFICE_PROCESSES = (
    # Standard Office applications
    "winword.exe",
    "excel.exe",
    "outlook.exe",
    "onenote.exe",
    "visio.exe",
    "powerpnt.exe",
    "msaccess.exe",
    "mspub.exe",
    "winproj.exe",
    "lync.exe",
    "teams.exe",
)
"""!
@brief Foreground Office applications terminated prior to uninstalls.
"""

OFFICE_PROCESS_PATTERNS = (
    "ose*.exe",
    "integrator.exe",
)
"""!
@brief Wildcard process filters used to mirror OffScrub cleanup loops.
"""

# Extended process list from VBS CloseOfficeApps (OffScrubC2R.vbs)
C2R_INFRASTRUCTURE_PROCESSES = (
    # C2R client and infrastructure
    "appvshnotify.exe",
    "integratedoffice.exe",
    "integrator.exe",
    "firstrun.exe",
    "officeclicktorun.exe",
    "officeondemand.exe",
    "OfficeC2RClient.exe",
    "AppVLP.exe",
    # Background processes
    "msosync.exe",
    "OneNoteM.exe",
    "communicator.exe",
    "perfboost.exe",
    "roamingoffice.exe",
    "mavinject32.exe",
    # Telemetry and update
    "OfficeTelemetryAgentLogOn.exe",
    "msouc.exe",
    "msoia.exe",
    "msoadfsb.exe",
    # Legacy processes
    "bcssync.exe",
    "officesas.exe",
    "officesasscheduler.exe",
    "groove.exe",
    "groovemonitor.exe",
    # Additional processes from OfficeScrubber.cmd/VBS (verified 2025)
    "werfault.exe",  # Windows Error Reporting - may interfere
    "mstore.exe",  # Microsoft Store related
    "setlang.exe",  # Language settings
    "ois.exe",  # Organization Info Service
    "graph.exe",  # Microsoft Graph
    "OfficeHubTaskHost.exe",  # Office Hub task
    "msoidsvc.exe",  # Office Identity Service
    "msoidsvcm.exe",  # Office Identity Service Manager
    "ucmapi.exe",  # Unified Communications MAPI
    "sdxhelper.exe",  # SDX Helper
    "OfficeClickToRun.exe",  # Alternate C2R name
    "officec2rclient.exe",  # Lowercase variant
)
"""!
@brief C2R infrastructure processes from VBS CloseOfficeApps.
@details These are background processes that should be terminated
before C2R uninstall or cleanup operations.
"""

MSI_INFRASTRUCTURE_PROCESSES = (
    # Office Source Engine
    "ose.exe",
    "ose64.exe",
    # Setup and maintenance
    "setup.exe",
    "msiexec.exe",
    "OfficeSetup.exe",
    # Shared components
    "msosync.exe",
    "groovemonitor.exe",
    "groove.exe",
)
"""!
@brief MSI infrastructure processes from VBS OffScrub scripts.
"""

ALL_OFFICE_PROCESSES = (
    *DEFAULT_OFFICE_PROCESSES,
    *C2R_INFRASTRUCTURE_PROCESSES,
    *MSI_INFRASTRUCTURE_PROCESSES,
)
"""!
@brief Combined list of all Office-related processes for full cleanup.
"""

MSI_UNINSTALL_ROOTS: tuple[tuple[int, str], ...] = (
    (HKLM, r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"),
    (HKLM, r"SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"),
)

OFFSCRUB_UNINSTALL_SEQUENCE: tuple[str, ...] = (
    "c2r",
    "2016",
    "2013",
    "2010",
    "2007",
    "2003",
)
"""!
@brief Ordered uninstall sequence mirrored from OfficeScrubber.cmd.
"""

OFFSCRUB_UNINSTALL_PRIORITY = {
    version: index for index, version in enumerate(OFFSCRUB_UNINSTALL_SEQUENCE)
}
"""!
@brief Priority lookup so planners can sort uninstall steps deterministically.
"""

MSI_UNINSTALL_VERSION_GROUPS: Mapping[str, str] = {
    "2024": "2016",
    "2021": "2016",
    "2019": "2016",
    "2016": "2016",
    "2013": "2013",
    "2010": "2010",
    "2007": "2007",
    "2003": "2003",
}
"""!
@brief Map MSI version identifiers to their OffScrub grouping.
"""

C2R_UNINSTALL_VERSION_GROUPS: Mapping[str, str] = {
    "365": "c2r",
    "2024": "c2r",
    "2021": "c2r",
    "2019": "c2r",
    "2016": "c2r",
}
"""!
@brief Map Click-to-Run version markers to the shared OffScrub stage.
"""

_VERSION_MAJOR_KEYS = ("11.0", "12.0", "14.0", "15.0", "16.0")


def _normalize_registry_entries(entries: Iterable[tuple[int, str]]) -> tuple[tuple[int, str], ...]:
    normalized: list[tuple[int, str]] = []
    seen: set[tuple[int, str]] = set()
    for hive, path in entries:
        canonical = path.replace("/", "\\").strip("\\")
        while "\\\\" in canonical:
            canonical = canonical.replace("\\\\", "\\")
        entry = (hive, canonical)
        if entry in seen:
            continue
        seen.add(entry)
        normalized.append(entry)
    return tuple(normalized)


def _registry_entry_depth(entry: tuple[int, str]) -> int:
    _, path = entry
    path = path.strip("\\")
    if not path:
        return 0
    return path.count("\\")


def _sort_registry_entries_deepest_first(
    entries: Iterable[tuple[int, str]],
) -> tuple[tuple[int, str], ...]:
    indexed = list(enumerate(entries))
    indexed.sort(key=lambda item: (-_registry_entry_depth(item[1]), item[0]))
    return tuple(entry for _, entry in indexed)


_REGISTRY_RESIDUE_BASE: list[tuple[int, str]] = [
    (HKLM, r"SOFTWARE\Microsoft\Office"),
    (HKLM, r"SOFTWARE\WOW6432Node\Microsoft\Office"),
    (HKCU, r"SOFTWARE\Microsoft\Office"),
    (HKCU, r"SOFTWARE\Policies\Microsoft\Office"),
    (HKLM, r"SOFTWARE\Policies\Microsoft\Office"),
    (HKLM, r"SOFTWARE\WOW6432Node\Policies\Microsoft\Office"),
    (HKCU, r"SOFTWARE\Policies\Microsoft\Cloud\Office"),
    (HKLM, r"SOFTWARE\Policies\Microsoft\Cloud\Office"),
    (HKLM, r"SOFTWARE\WOW6432Node\Policies\Microsoft\Cloud\Office"),
    (HKLM, r"SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"),
    (HKLM, r"SOFTWARE\Microsoft\Office\ClickToRun"),
    (HKLM, r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"),
    (HKLM, r"SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs"),
    (HKLM, r"SOFTWARE\Microsoft\Office\ClickToRun\Updates"),
    (
        HKLM,
        r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
        r"\0ff1ce15-a989-479d-af46-f275c6370663",
    ),
    (
        HKLM,
        r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\Policies"
        r"\0ff1ce15-a989-479d-af46-f275c6370663",
    ),
    (
        HKU,
        r"S-1-5-20\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
        r"\0ff1ce15-a989-479d-af46-f275c6370663",
    ),
    (
        HKU,
        r"S-1-5-20\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\Policies"
        r"\0ff1ce15-a989-479d-af46-f275c6370663",
    ),
    (
        HKLM,
        r"SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663",
    ),
    (
        HKLM,
        r"SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform\59a52881-a989-479d-af46-f275c6370663",
    ),
    (HKU, r"S-1-5-20\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"),
    (HKLM, r"SOFTWARE\Microsoft\Office\16.0\Common\OEM"),
    (HKLM, r"SOFTWARE\Microsoft\Office\16.0\Common\Licensing"),
    (HKLM, r"SOFTWARE\Policies\Microsoft\Office\16.0\Common\Licensing"),
    (HKLM, r"SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Common\Licensing"),
    (HKLM, r"SOFTWARE\WOW6432Node\Policies\Microsoft\Office\16.0\Common\Licensing"),
]

for major in _VERSION_MAJOR_KEYS:
    _REGISTRY_RESIDUE_BASE.append((HKCU, rf"SOFTWARE\Microsoft\Office\{major}"))
    _REGISTRY_RESIDUE_BASE.append((HKLM, rf"SOFTWARE\Microsoft\Office\{major}"))
    _REGISTRY_RESIDUE_BASE.append((HKLM, rf"SOFTWARE\WOW6432Node\Microsoft\Office\{major}"))
    _REGISTRY_RESIDUE_BASE.append(
        (HKLM, rf"SOFTWARE\WOW6432Node\Policies\Microsoft\Office\{major}")
    )
    _REGISTRY_RESIDUE_BASE.append(
        (HKLM, rf"SOFTWARE\WOW6432Node\Policies\Microsoft\Cloud\Office\{major}")
    )

# COM/CLSIDs tied to Office components (e.g., MAPI search handlers)
_OFFICE_COM_CLSIDS = (
    "{2027FC3B-CF9D-4EC7-A823-38BA308625CC}",
    "{573FFD05-2805-47C2-BCE0-5F19512BEB8D}",
    "{8BA85C75-763B-4103-94EB-9470F12FE0F7}",
    "{CD55129A-B1A1-438E-A425-CEBC7DC684EE}",
    "{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}",
    "{E768CD3B-BDDC-436D-9C13-E1B39CA257B1}",
    "{F8E61EDD-EA25-484E-AC8A-7447F2AAE2A9}",
)

for clsid in _OFFICE_COM_CLSIDS:
    _REGISTRY_RESIDUE_BASE.append((HKLM, rf"SOFTWARE\Classes\CLSID\{clsid}"))
    _REGISTRY_RESIDUE_BASE.append((HKLM, rf"SOFTWARE\WOW6432Node\Classes\CLSID\{clsid}"))

# Shell Integration paths (from VBS RegWipe - protocol handlers, context menu, etc.)
_SHELL_INTEGRATION_PATHS: list[tuple[int, str]] = [
    # Protocol handlers (osf:, onenote:, etc.)
    (HKLM, r"SOFTWARE\Classes\osf"),
    (HKLM, r"SOFTWARE\Classes\osf-roaming.16"),
    (HKLM, r"SOFTWARE\Classes\onenote"),
    (HKLM, r"SOFTWARE\Classes\onenote-cmd"),
    (HKLM, r"SOFTWARE\Classes\onenote15"),
    (HKLM, r"SOFTWARE\Classes\OneNote.URL.16"),
    (HKLM, r"SOFTWARE\Classes\ms-word"),
    (HKLM, r"SOFTWARE\Classes\ms-excel"),
    (HKLM, r"SOFTWARE\Classes\ms-powerpoint"),
    (HKLM, r"SOFTWARE\Classes\ms-access"),
    (HKLM, r"SOFTWARE\Classes\ms-publisher"),
    (HKLM, r"SOFTWARE\Classes\ms-outlook"),
    (HKLM, r"SOFTWARE\Classes\ms-visio"),
    (HKLM, r"SOFTWARE\Classes\ms-project"),
    (HKLM, r"SOFTWARE\Classes\ms-spd"),
    (HKCU, r"SOFTWARE\Classes\osf"),
    (HKCU, r"SOFTWARE\Classes\osf-roaming.16"),
    # Shell icon overlays (OneDrive, Groove sync overlays)
    (
        HKLM,
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\GDriveSyncedOverlay",
    ),
    (
        HKLM,
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\GDriveSyncingOverlay",
    ),
    (
        HKLM,
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\GDrivePausedOverlay",
    ),
    (
        HKLM,
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\GDriveWarningOverlay",
    ),
    (
        HKLM,
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\GDriveErrorOverlay",
    ),
    # Explorer namespace extensions
    (
        HKLM,
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{018D5C66-4533-4307-9B53-224DE2ED1FE6}",
    ),
    (
        HKLM,
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{018D5C66-4533-4307-9B53-224DE2ED1FE6}",
    ),
    # Browser helper objects
    (
        HKLM,
        # BHO registration for Office document handling
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer"
        r"\Browser Helper Objects\{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}",
    ),
    (
        HKLM,
        # BHO registration for Office document handling (32-bit)
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Explorer"
        r"\Browser Helper Objects\{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}",
    ),
    # Shell approved extensions
    (
        HKLM,
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved"
        r"\{42042206-2D85-11D3-8CFF-005004838597}",
    ),
    (
        HKLM,
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved"
        r"\{993BE281-6695-4BA5-8A2A-7AACBFAAB69E}",
    ),
    (
        HKLM,
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved"
        r"\{0006F045-0000-0000-C000-000000000046}",
    ),
    # Context menu handlers
    (HKLM, r"SOFTWARE\Classes\*\shellex\ContextMenuHandlers\GrooveShellExtensions"),
    (HKLM, r"SOFTWARE\Classes\Directory\shellex\ContextMenuHandlers\GrooveShellExtensions"),
    (HKLM, r"SOFTWARE\Classes\Folder\shellex\ContextMenuHandlers\GrooveShellExtensions"),
    # Default mail client registration (Control Panel "Mail (Microsoft Outlook)")
    (HKLM, r"SOFTWARE\Clients\Mail\Microsoft Outlook"),
    (HKLM, r"SOFTWARE\WOW6432Node\Clients\Mail\Microsoft Outlook"),
    # Mail Control Panel applet registration (CLSID {A0D4CD32-5D5D-4f72-BAAA-767A7AD6BAC5})
    (HKLM, r"SOFTWARE\Classes\CLSID\{A0D4CD32-5D5D-4f72-BAAA-767A7AD6BAC5}"),
    (HKLM, r"SOFTWARE\WOW6432Node\Classes\CLSID\{A0D4CD32-5D5D-4f72-BAAA-767A7AD6BAC5}"),
    (
        HKLM,
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel\NameSpace\{A0D4CD32-5D5D-4f72-BAAA-767A7AD6BAC5}",
    ),
]
_REGISTRY_RESIDUE_BASE.extend(_SHELL_INTEGRATION_PATHS)

# Windows Installer metadata paths (from VBS ValidateWIMetadataKey)
_WI_METADATA_PATTERNS: list[tuple[int, str]] = [
    # Note: Actual cleanup requires product code enumeration, these are base paths
    (HKLM, r"SOFTWARE\Classes\Installer\Products"),
    (HKLM, r"SOFTWARE\Classes\Installer\Features"),
    (HKLM, r"SOFTWARE\Classes\Installer\Components"),
    (HKLM, r"SOFTWARE\Classes\Installer\Patches"),
    (HKLM, r"SOFTWARE\Classes\Installer\UpgradeCodes"),
    (HKLM, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData"),
]
# Don't add to residue base - these need special handling via msi_components module

# TypeLib GUIDs for Office (from VBS OffScrub TypeLib cleanup)
# These are cleaned if the referenced DLL no longer exists
_OFFICE_TYPELIB_GUIDS = (
    "{00020430-0000-0000-C000-000000000046}",  # StdOle
    "{00020802-0000-0000-C000-000000000046}",  # MS Forms
    "{00020813-0000-0000-C000-000000000046}",  # Excel
    "{00020905-0000-0000-C000-000000000046}",  # Word
    "{00020970-0000-0000-C000-000000000046}",  # Word (alternate)
    "{000204EF-0000-0000-C000-000000000046}",  # Outlook
    "{00021A20-0000-0000-C000-000000000046}",  # Outlook (Forms)
    "{00024500-0000-0000-C000-000000000046}",  # PowerPoint
    "{0002E157-0000-0000-C000-000000000046}",  # VBE
    "{0002E55D-0000-0000-C000-000000000046}",  # VBE Objects
    "{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}",  # Office Drawing
    "{0D452EE1-E08F-101A-852E-02608C4D0BB4}",  # VBAEN
    "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}",  # SHDOCVW (Shell Doc Object)
    "{91493440-5A91-11CF-8700-00AA0060263B}",  # PowerPoint (Full)
    "{AC0714F2-3D04-11D1-AE7D-00A0C90F26F4}",  # VBComponents
    "{EF53050B-882E-4776-B643-EDA472E8E3F2}",  # MSOXMLMF
    "{00062FFF-0000-0000-C000-000000000046}",  # Outlook Object Model
)

# Add-in registration paths (from VBS /CLEARADDINREG)
_ADDIN_REGISTRY_PATHS: list[tuple[int, str]] = [
    (HKCU, r"SOFTWARE\Microsoft\Office\Excel\AddIns"),
    (HKCU, r"SOFTWARE\Microsoft\Office\Word\AddIns"),
    (HKCU, r"SOFTWARE\Microsoft\Office\Outlook\AddIns"),
    (HKCU, r"SOFTWARE\Microsoft\Office\PowerPoint\AddIns"),
    (HKCU, r"SOFTWARE\Microsoft\Office\Access\AddIns"),
    (HKCU, r"SOFTWARE\Microsoft\Office\Publisher\AddIns"),
    (HKCU, r"SOFTWARE\Microsoft\Office\Visio\AddIns"),
    (HKCU, r"SOFTWARE\Microsoft\Office\Project\AddIns"),
    (HKCU, r"SOFTWARE\Microsoft\Office\OneNote\AddIns"),
    (HKLM, r"SOFTWARE\Microsoft\Office\Excel\AddIns"),
    (HKLM, r"SOFTWARE\Microsoft\Office\Word\AddIns"),
    (HKLM, r"SOFTWARE\Microsoft\Office\Outlook\AddIns"),
    (HKLM, r"SOFTWARE\Microsoft\Office\PowerPoint\AddIns"),
    (HKLM, r"SOFTWARE\Microsoft\Office\Access\AddIns"),
    (HKLM, r"SOFTWARE\Microsoft\Office\Publisher\AddIns"),
    (HKLM, r"SOFTWARE\Microsoft\Office\Visio\AddIns"),
    (HKLM, r"SOFTWARE\Microsoft\Office\Project\AddIns"),
    (HKLM, r"SOFTWARE\Microsoft\Office\OneNote\AddIns"),
    (HKLM, r"SOFTWARE\WOW6432Node\Microsoft\Office\Excel\AddIns"),
    (HKLM, r"SOFTWARE\WOW6432Node\Microsoft\Office\Word\AddIns"),
    (HKLM, r"SOFTWARE\WOW6432Node\Microsoft\Office\Outlook\AddIns"),
    (HKLM, r"SOFTWARE\WOW6432Node\Microsoft\Office\PowerPoint\AddIns"),
]
_REGISTRY_RESIDUE_BASE.extend(_ADDIN_REGISTRY_PATHS)

# Scheduled Tasks and Services registry (from VBS /OSE)
_SERVICE_REGISTRY_PATHS: list[tuple[int, str]] = [
    (HKLM, r"SYSTEM\CurrentControlSet\Services\ClickToRunSvc"),
    (HKLM, r"SYSTEM\CurrentControlSet\Services\OfficeSvc"),
    (HKLM, r"SYSTEM\CurrentControlSet\Services\ose"),
    (HKLM, r"SYSTEM\CurrentControlSet\Services\ose64"),
    (HKLM, r"SYSTEM\CurrentControlSet\Services\osppsvc"),
    (HKLM, r"SYSTEM\CurrentControlSet\Services\osppobject"),
]
# Services are not deleted via registry - handled by tasks_services module

# Office service names for deletion (from VBS DeleteService)
OFFICE_SERVICES_TO_DELETE: tuple[str, ...] = (
    "ClickToRunSvc",  # C2R main service
    "OfficeSvc",  # Office service
    "ose",  # Office Source Engine (32-bit)
    "ose64",  # Office Source Engine (64-bit)
    "osppsvc",  # Office Software Protection Platform
)
"""!
@brief Service names to stop and delete during Office removal.
@details Based on VBS DeleteService calls in OffScrubC2R.vbs.
"""

# Scheduled task names for deletion (from VBS DelSchtasks in OffScrubC2R.vbs)
OFFICE_SCHEDULED_TASKS_TO_DELETE: tuple[str, ...] = (
    # C2R streaming and update tasks
    "FF_INTEGRATEDstreamSchedule",
    "FF_INTEGRATEDUPDATEDETECTION",
    "C2RAppVLoggingStart",
    # Subscription and heartbeat tasks
    "Office 15 Subscription Heartbeat",
    r"Microsoft Office 15 Sync Maintenance for {d068b555-9700-40b8-992c-f866287b06c1}",
    "Office Subscription Maintenance",
    "Office Background Streaming",
    # Telemetry and inventory tasks (under \Microsoft\Office\)
    r"\Microsoft\Office\OfficeInventoryAgentFallBack",
    r"\Microsoft\Office\OfficeTelemetryAgentFallBack",
    r"\Microsoft\Office\OfficeInventoryAgentLogOn",
    r"\Microsoft\Office\OfficeTelemetryAgentLogOn",
    # C2R service monitoring tasks
    r"\Microsoft\Office\Office Automatic Updates",
    r"\Microsoft\Office\Office ClickToRun Service Monitor",
)
"""!
@brief Scheduled task names to delete during Office removal.
@details Based on VBS DelSchtasks subroutine in OffScrubC2R.vbs (lines 1053-1108).
"""

# Msiexec return codes from VBS SetupRetVal (OffScrubC2R.vbs lines 3306-3362)
MSIEXEC_RETURN_CODES: dict[int, str] = {
    0: "SUCCESS",
    1259: "APPHELP_BLOCK",
    1601: "INSTALL_SERVICE_FAILURE",
    1602: "INSTALL_USEREXIT",
    1603: "INSTALL_FAILURE",
    1604: "INSTALL_SUSPEND",
    1605: "UNKNOWN_PRODUCT",
    1606: "UNKNOWN_FEATURE",
    1607: "UNKNOWN_COMPONENT",
    1608: "UNKNOWN_PROPERTY",
    1609: "INVALID_HANDLE_STATE",
    1610: "BAD_CONFIGURATION",
    1611: "INDEX_ABSENT",
    1612: "INSTALL_SOURCE_ABSENT",
    1613: "INSTALL_PACKAGE_VERSION",
    1614: "PRODUCT_UNINSTALLED",
    1615: "BAD_QUERY_SYNTAX",
    1616: "INVALID_FIELD",
    1618: "INSTALL_ALREADY_RUNNING",
    1619: "INSTALL_PACKAGE_OPEN_FAILED",
    1620: "INSTALL_PACKAGE_INVALID",
    1621: "INSTALL_UI_FAILURE",
    1622: "INSTALL_LOG_FAILURE",
    1623: "INSTALL_LANGUAGE_UNSUPPORTED",
    1624: "INSTALL_TRANSFORM_FAILURE",
    1625: "INSTALL_PACKAGE_REJECTED",
    1626: "FUNCTION_NOT_CALLED",
    1627: "FUNCTION_FAILED",
    1628: "INVALID_TABLE",
    1629: "DATATYPE_MISMATCH",
    1630: "UNSUPPORTED_TYPE",
    1631: "CREATE_FAILED",
    1632: "INSTALL_TEMP_UNWRITABLE",
    1633: "INSTALL_PLATFORM_UNSUPPORTED",
    1634: "INSTALL_NOTUSED",
    1635: "PATCH_PACKAGE_OPEN_FAILED",
    1636: "PATCH_PACKAGE_INVALID",
    1637: "PATCH_PACKAGE_UNSUPPORTED",
    1638: "PRODUCT_VERSION",
    1639: "INVALID_COMMAND_LINE",
    1640: "INSTALL_REMOTE_DISALLOWED",
    1641: "SUCCESS_REBOOT_INITIATED",
    1642: "PATCH_TARGET_NOT_FOUND",
    1643: "PATCH_PACKAGE_REJECTED",
    1644: "INSTALL_TRANSFORM_REJECTED",
    1645: "INSTALL_REMOTE_PROHIBITED",
    1646: "PATCH_REMOVAL_UNSUPPORTED",
    1647: "UNKNOWN_PATCH",
    1648: "PATCH_NO_SEQUENCE",
    1649: "PATCH_REMOVAL_DISALLOWED",
    1650: "INVALID_PATCH_XML",
    3010: "SUCCESS_REBOOT_REQUIRED",
}
"""!
@brief Msiexec return code translation table.
@details Based on VBS SetupRetVal function in OffScrubC2R.vbs (lines 3306-3362).
"""


def translate_msiexec_return_code(code: int) -> str:
    """!
    @brief Translate an msiexec return code to a human-readable string.
    @param code The return code from msiexec.exe.
    @returns String description of the code, or "UNKNOWN" if not recognized.
    """
    return MSIEXEC_RETURN_CODES.get(code, f"UNKNOWN ({code})")


# C2R-specific residue paths (from OffScrubC2R.vbs RegWipe)
_C2R_REGISTRY_RESIDUE: list[tuple[int, str]] = [
    (HKLM, r"SOFTWARE\Microsoft\Office\ClickToRun"),
    (HKLM, r"SOFTWARE\Microsoft\Office\15.0\ClickToRun"),
    (HKLM, r"SOFTWARE\WOW6432Node\Microsoft\Office\15.0\ClickToRun"),
    (HKLM, r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"),
    (HKLM, r"SOFTWARE\Microsoft\Office\ClickToRun\PackageGUID"),
    (HKLM, r"SOFTWARE\Microsoft\Office\ClickToRun\PackageRoot"),
    (HKLM, r"SOFTWARE\Microsoft\Office\ClickToRun\Platform"),
    (HKLM, r"SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs"),
    (HKLM, r"SOFTWARE\Microsoft\Office\ClickToRun\PropertyBag"),
    (HKLM, r"SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY"),
    (HKLM, r"SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE"),
    (HKLM, r"SOFTWARE\Microsoft\Office\ClickToRun\Scenario"),
    (HKLM, r"SOFTWARE\Microsoft\Office\ClickToRun\Updates"),
    (HKCU, r"SOFTWARE\Microsoft\Office\ClickToRun"),
]
_REGISTRY_RESIDUE_BASE.extend(_C2R_REGISTRY_RESIDUE)

REGISTRY_RESIDUE_PATHS = _sort_registry_entries_deepest_first(
    _normalize_registry_entries(_REGISTRY_RESIDUE_BASE)
)
"""!
@brief Registry residue handles derived from OfficeScrubber cleanup routines.
"""

RESIDUE_PATH_TEMPLATES = (
    # Licensing directories
    {
        "label": "programdata_office_licenses",
        "path": r"%PROGRAMDATA%\\Microsoft\\Office\\Licenses",
        "category": "licenses",
    },
    {
        "label": "programdata_microsoft_licenses",
        "path": r"%PROGRAMDATA%\\Microsoft\\Licenses",
        "category": "licenses",
    },
    {
        "label": "localappdata_office_licenses",
        "path": r"%LOCALAPPDATA%\\Microsoft\\Office\\Licenses",
        "category": "licenses",
    },
    {
        "label": "localappdata_office_licensing16",
        "path": r"%LOCALAPPDATA%\\Microsoft\\Office\\16.0\\Licensing",
        "category": "licenses",
    },
    # Identity and authentication caches
    {
        "label": "localappdata_identity_cache",
        "path": r"%LOCALAPPDATA%\\Microsoft\\IdentityCache",
        "category": "identity",
    },
    {
        "label": "localappdata_oneauth",
        "path": r"%LOCALAPPDATA%\\Microsoft\\OneAuth",
        "category": "identity",
    },
    # Click-to-Run cache and installation directories
    {
        "label": "programdata_clicktorun",
        "path": r"%PROGRAMDATA%\\Microsoft\\ClickToRun",
        "category": "c2r_cache",
    },
    {
        "label": "programfiles_clicktorun_x86",
        "path": r"C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\ClickToRun",
        "category": "c2r_cache",
    },
    {
        "label": "programfiles_clicktorun_x64",
        "path": r"C:\\Program Files\\Common Files\\Microsoft Shared\\ClickToRun",
        "category": "c2r_cache",
    },
    # ProgramData Office directories (left behind after partial uninstall)
    {
        "label": "programdata_office",
        "path": r"%PROGRAMDATA%\\Microsoft\\Office",
        "category": "office_data",
    },
    # Windows Installer cache (common leftover)
    {
        "label": "windows_installer_office_cache_x86",
        "path": r"C:\\Windows\\Installer\\{90160000-",
        "category": "installer_cache",
    },
    # VBA directories
    {
        "label": "appdata_vba",
        "path": r"%APPDATA%\\Microsoft\\VBA",
        "category": "vba",
    },
    {
        "label": "localappdata_vba",
        "path": r"%LOCALAPPDATA%\\Microsoft\\VBA",
        "category": "vba",
    },
    # Office Setup bootstrap and update directories
    {
        "label": "programdata_office_updates",
        "path": r"%PROGRAMDATA%\\Microsoft\\Office\\Updates",
        "category": "updates",
    },
    {
        "label": "programfiles_office_updates_x86",
        "path": r"C:\\Program Files (x86)\\Microsoft Office 15",
        "architecture": "x86",
        "category": "legacy_office",
    },
    {
        "label": "programfiles_office_updates_x64",
        "path": r"C:\\Program Files\\Microsoft Office 15",
        "architecture": "x64",
        "category": "legacy_office",
    },
    # C2R Package folders (from OffScrubC2R.vbs FolderWipe)
    {
        "label": "programfiles_office16_root_x64",
        "path": r"C:\\Program Files\\Microsoft Office\\root",
        "architecture": "x64",
        "category": "c2r_install",
    },
    {
        "label": "programfiles_office16_root_x86",
        "path": r"C:\\Program Files (x86)\\Microsoft Office\\root",
        "architecture": "x86",
        "category": "c2r_install",
    },
    {
        "label": "programfiles_office16_x64",
        "path": r"C:\\Program Files\\Microsoft Office 16",
        "architecture": "x64",
        "category": "c2r_install",
    },
    {
        "label": "programfiles_office16_x86",
        "path": r"C:\\Program Files (x86)\\Microsoft Office 16",
        "architecture": "x86",
        "category": "c2r_install",
    },
    {
        "label": "programfiles_office15_x64",
        "path": r"C:\\Program Files\\Microsoft Office 15",
        "architecture": "x64",
        "category": "c2r_install",
    },
    {
        "label": "programfiles_office15_x86",
        "path": r"C:\\Program Files (x86)\\Microsoft Office 15",
        "architecture": "x86",
        "category": "c2r_install",
    },
    # Common Files shared components
    {
        "label": "commonfiles_office16_x64",
        "path": r"C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE16",
        "architecture": "x64",
        "category": "shared_components",
    },
    {
        "label": "commonfiles_office16_x86",
        "path": r"C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\OFFICE16",
        "architecture": "x86",
        "category": "shared_components",
    },
    {
        "label": "commonfiles_office15_x64",
        "path": r"C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE15",
        "architecture": "x64",
        "category": "shared_components",
    },
    {
        "label": "commonfiles_office15_x86",
        "path": r"C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\OFFICE15",
        "architecture": "x86",
        "category": "shared_components",
    },
    {
        "label": "commonfiles_vba_x64",
        "path": r"C:\\Program Files\\Common Files\\Microsoft Shared\\VBA",
        "architecture": "x64",
        "category": "shared_components",
    },
    {
        "label": "commonfiles_vba_x86",
        "path": r"C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\VBA",
        "architecture": "x86",
        "category": "shared_components",
    },
    # MSOCache on system drives (from VBS CleanOffice.txt)
    {
        "label": "msocache_system",
        "path": r"C:\\MSOCache",
        "category": "installer_cache",
    },
    # User Office data directories
    {
        "label": "appdata_office",
        "path": r"%APPDATA%\\Microsoft\\Office",
        "category": "user_data",
    },
    {
        "label": "localappdata_office",
        "path": r"%LOCALAPPDATA%\\Microsoft\\Office",
        "category": "user_data",
    },
    {
        "label": "appdata_templates",
        "path": r"%APPDATA%\\Microsoft\\Templates",
        "category": "user_data",
    },
    # Document Recovery
    {
        "label": "localappdata_word_recovery",
        "path": r"%LOCALAPPDATA%\\Microsoft\\Word",
        "category": "recovery",
    },
    {
        "label": "localappdata_excel_recovery",
        "path": r"%LOCALAPPDATA%\\Microsoft\\Excel",
        "category": "recovery",
    },
    {
        "label": "localappdata_powerpoint_recovery",
        "path": r"%LOCALAPPDATA%\\Microsoft\\PowerPoint",
        "category": "recovery",
    },
    # PackageManifests (C2R)
    {
        "label": "programdata_packagemanifests",
        "path": r"%PROGRAMDATA%\\Microsoft\\Office\\PackageManifests",
        "category": "c2r_cache",
    },
)
"""!
@brief Filesystem residue directories removed by the reference scripts.
"""

# Export the TypeLib GUIDs for use by registry cleanup
OFFICE_TYPELIB_GUIDS = _OFFICE_TYPELIB_GUIDS
"""!
@brief TypeLib GUIDs that may reference Office DLLs.
@details Used by registry cleanup to detect orphaned TypeLib registrations.
"""


def _merge_roots(*roots: tuple[int, str]) -> tuple[tuple[int, str], ...]:
    """!
    @brief Provide a helper that returns a deduplicated tuple of uninstall roots.
    """

    seen: set[tuple[int, str]] = set()
    ordered: list[tuple[int, str]] = []
    for hive, path in roots:
        entry = (hive, path)
        if entry in seen:
            continue
        ordered.append(entry)
        seen.add(entry)
    return tuple(ordered)


OFFSCRUB_EXECUTABLE = "cscript.exe"
"""!
@brief Host executable used by OffScrub VBS helpers.
"""

OFFSCRUB_HOST_ARGS: tuple[str, ...] = ("//NoLogo",)
"""!
@brief Common arguments prepended to every OffScrub invocation.
"""

MSI_OFFSCRUB_SCRIPT_MAP: dict[str, str] = {
    "2003": "OffScrub03.vbs",
    "2007": "OffScrub07.vbs",
    "2010": "OffScrub10.vbs",
    "2013": "OffScrub_O15msi.vbs",
    "2016": "OffScrub_O16msi.vbs",
    "2019": "OffScrub_O16msi.vbs",
    "2021": "OffScrub_O16msi.vbs",
    "2024": "OffScrub_O16msi.vbs",
    "365": "OffScrub_O16msi.vbs",
}
"""!
@brief Mapping between Office versions and MSI OffScrub helpers.
"""

MSI_OFFSCRUB_DEFAULT_SCRIPT = "OffScrub_O16msi.vbs"
"""!
@brief Default MSI OffScrub helper when no specific version matches.
"""

MSI_OFFSCRUB_ARGS: tuple[str, ...] = (
    "ALL",
    "/OSE",
    "/NOCANCEL",
    "/FORCE",
    "/ENDCURRENTINSTALLS",
    "/DELETEUSERSETTINGS",
    "/CLEARADDINREG",
    "/REMOVELYNC",
)
"""!
@brief Argument list mirrored from OfficeScrubber MSI automation.
"""

C2R_OFFSCRUB_SCRIPT = "OffScrubC2R.vbs"
"""!
@brief Click-to-Run OffScrub helper name mirrored from OfficeScrubber.
"""

C2R_OFFSCRUB_ARGS: tuple[str, ...] = ("ALL", "/OFFLINE")
"""!
@brief Arguments passed to the Click-to-Run OffScrub helper.
"""

UNINSTALL_COMMAND_TEMPLATES: dict[str, dict[str, object]] = {
    "msi": {
        "executable": OFFSCRUB_EXECUTABLE,
        "host_args": OFFSCRUB_HOST_ARGS,
        "script_map": MSI_OFFSCRUB_SCRIPT_MAP,
        "default_script": MSI_OFFSCRUB_DEFAULT_SCRIPT,
        "arguments": MSI_OFFSCRUB_ARGS,
    },
    "c2r": {
        "executable": OFFSCRUB_EXECUTABLE,
        "host_args": OFFSCRUB_HOST_ARGS,
        "script": C2R_OFFSCRUB_SCRIPT,
        "arguments": C2R_OFFSCRUB_ARGS,
    },
}
"""!
@brief Templates describing how MSI and C2R uninstalls are invoked.
"""


MSI_PRODUCT_MAP: dict[str, dict[str, object]] = {
    # Office 2010 (14.x) Professional Plus
    "{90140000-0011-0000-0000-0000000FF1CE}": {
        "product": "Microsoft Office Professional Plus 2010",
        "edition": "Professional Plus",
        "version": "2010",
        "supported_versions": ("2010",),
        "architecture": "x86",
        "registry_roots": MSI_UNINSTALL_ROOTS,
        "family": "office",
    },
    "{90140000-0011-0000-1000-0000000FF1CE}": {
        "product": "Microsoft Office Professional Plus 2010",
        "edition": "Professional Plus",
        "version": "2010",
        "supported_versions": ("2010",),
        "architecture": "x64",
        "registry_roots": _merge_roots(MSI_UNINSTALL_ROOTS[0]),
        "family": "office",
    },
    # Office 2013 (15.x) Professional Plus
    "{90150000-0011-0000-0000-0000000FF1CE}": {
        "product": "Microsoft Office Professional Plus 2013",
        "edition": "Professional Plus",
        "version": "2013",
        "supported_versions": ("2013",),
        "architecture": "x86",
        "registry_roots": MSI_UNINSTALL_ROOTS,
        "family": "office",
    },
    "{90150000-0011-0000-1000-0000000FF1CE}": {
        "product": "Microsoft Office Professional Plus 2013",
        "edition": "Professional Plus",
        "version": "2013",
        "supported_versions": ("2013",),
        "architecture": "x64",
        "registry_roots": _merge_roots(MSI_UNINSTALL_ROOTS[0]),
        "family": "office",
    },
    # Office 2016/2019/2021/2024 perpetual channel (MSI-based SKUs)
    "{90160000-0011-0000-0000-0000000FF1CE}": {
        "product": "Microsoft Office Professional Plus 2016",
        "edition": "Professional Plus",
        "version": "2016",
        "supported_versions": ("2016", "2019", "2021", "2024"),
        "architecture": "x86",
        "registry_roots": MSI_UNINSTALL_ROOTS,
        "family": "office",
    },
    "{90160000-0011-0000-1000-0000000FF1CE}": {
        "product": "Microsoft Office Professional Plus 2016",
        "edition": "Professional Plus",
        "version": "2016",
        "supported_versions": ("2016", "2019", "2021", "2024"),
        "architecture": "x64",
        "registry_roots": _merge_roots(MSI_UNINSTALL_ROOTS[0]),
        "family": "office",
    },
    # Visio Professional 2016 (MSI)
    "{90160000-0051-0000-0000-0000000FF1CE}": {
        "product": "Microsoft Visio Professional 2016",
        "edition": "Visio Professional",
        "version": "2016",
        "supported_versions": ("2016", "2019", "2021", "2024"),
        "architecture": "x86",
        "registry_roots": MSI_UNINSTALL_ROOTS,
        "family": "visio",
    },
    "{90160000-0051-0000-1000-0000000FF1CE}": {
        "product": "Microsoft Visio Professional 2016",
        "edition": "Visio Professional",
        "version": "2016",
        "supported_versions": ("2016", "2019", "2021", "2024"),
        "architecture": "x64",
        "registry_roots": _merge_roots(MSI_UNINSTALL_ROOTS[0]),
        "family": "visio",
    },
    # Project Professional 2016 (MSI)
    "{90160000-003B-0000-0000-0000000FF1CE}": {
        "product": "Microsoft Project Professional 2016",
        "edition": "Project Professional",
        "version": "2016",
        "supported_versions": ("2016", "2019", "2021", "2024"),
        "architecture": "x86",
        "registry_roots": MSI_UNINSTALL_ROOTS,
        "family": "project",
    },
    "{90160000-003B-0000-1000-0000000FF1CE}": {
        "product": "Microsoft Project Professional 2016",
        "edition": "Project Professional",
        "version": "2016",
        "supported_versions": ("2016", "2019", "2021", "2024"),
        "architecture": "x64",
        "registry_roots": _merge_roots(MSI_UNINSTALL_ROOTS[0]),
        "family": "project",
    },
}


def known_msi_codes() -> Iterable[str]:
    """!
    @brief Iterate the product codes that map to Office MSI deployments.
    """

    return MSI_PRODUCT_MAP.keys()


C2R_CONFIGURATION_KEYS: tuple[tuple[int, str], ...] = (
    (HKLM, r"SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration"),
    (HKLM, r"SOFTWARE\\Microsoft\\Office\\15.0\\ClickToRun\\Configuration"),
    (HKLM, r"SOFTWARE\\WOW6432Node\\Microsoft\\Office\\15.0\\ClickToRun\\Configuration"),
    (HKCU, r"SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration"),
)

C2R_PRODUCT_RELEASE_ROOTS: tuple[tuple[int, str], ...] = (
    (HKLM, r"SOFTWARE\\Microsoft\\Office\\ClickToRun\\ProductReleaseIDs"),
    (HKLM, r"SOFTWARE\\Microsoft\\Office\\15.0\\ClickToRun\\ProductReleaseIDs"),
)

C2R_SUBSCRIPTION_ROOTS: tuple[tuple[int, str], ...] = (
    (HKLM, r"SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration\\Subscriptions"),
    (HKCU, r"SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration\\Subscriptions"),
)

C2R_COM_REGISTRY_PATHS: tuple[tuple[int, str], ...] = (
    (HKLM, r"SOFTWARE\\Microsoft\\Office\\ClickToRun\\COM Compatibility\\Applications"),
    (HKCU, r"SOFTWARE\\Microsoft\\Office\\ClickToRun\\COM Compatibility\\Applications"),
)

C2R_PLATFORM_ALIASES: dict[str, str] = {
    "x86": "x86",
    "x64": "x64",
    "amd64": "x64",
    "arm64": "ARM64",
    "neutral": "neutral",
}

C2R_CHANNEL_ALIASES: dict[str, str] = {
    "Production::CC": "Current Channel",
    "Production::Current": "Current Channel",
    "Production::MEC": "Monthly Enterprise Channel",
    "Production::SAEC": "Semi-Annual Enterprise Channel",
    "Production::SA": "Semi-Annual Channel",
    "Production::Beta": "Beta Channel",
    "Production::InsiderFast": "Insider Fast",
    "Production::FirstReleaseCurrent": "Current Channel (Preview)",
    "Production::FirstReleaseDeferred": "Semi-Annual Preview",
    ("http://officecdn.microsoft.com/pr/492350f6-3a04-4b59-8b34-4c547755c2a0"): "Current Channel",
    (
        "http://officecdn.microsoft.com/pr/55336b82-a18d-4dd6-b5f6-9e5095c314a6"
    ): "Monthly Enterprise Channel",
    (
        "http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114"
    ): "Semi-Annual Enterprise Channel",
    (
        "http://officecdn.microsoft.com/pr/5440fd1f-7ecb-4221-8110-14e4edeeb5d0"
    ): "Semi-Annual Preview",
}

C2R_PRODUCT_RELEASES: Mapping[str, dict[str, object]] = {
    # Microsoft 365 Apps / Office 2016+ suites
    "O365ProPlusRetail": {
        "product": "Microsoft 365 Apps for enterprise",
        "supported_versions": ("2016", "2019", "2021", "2024", "365"),
        "architectures": ("x86", "x64", "ARM64"),
        "family": "office",
        "registry_paths": {
            "configuration": C2R_CONFIGURATION_KEYS,
            "product_release_ids": C2R_PRODUCT_RELEASE_ROOTS,
        },
    },
    "O365ProPlusVolume": {
        "product": "Microsoft 365 Apps for enterprise (Volume)",
        "supported_versions": ("2016", "2019", "2021", "2024", "365"),
        "architectures": ("x86", "x64", "ARM64"),
        "family": "office",
        "registry_paths": {
            "configuration": C2R_CONFIGURATION_KEYS,
            "product_release_ids": C2R_PRODUCT_RELEASE_ROOTS,
        },
    },
    "O365BusinessRetail": {
        "product": "Microsoft 365 Apps for business",
        "supported_versions": ("2016", "2019", "2021", "2024", "365"),
        "architectures": ("x86", "x64", "ARM64"),
        "family": "office",
        "registry_paths": {
            "configuration": C2R_CONFIGURATION_KEYS,
            "product_release_ids": C2R_PRODUCT_RELEASE_ROOTS,
        },
    },
    # Perpetual retail SKUs
    "ProPlus2019Retail": {
        "product": "Office Professional Plus 2019 (C2R)",
        "supported_versions": ("2019",),
        "architectures": ("x86", "x64"),
        "family": "office",
        "registry_paths": {
            "configuration": C2R_CONFIGURATION_KEYS,
            "product_release_ids": C2R_PRODUCT_RELEASE_ROOTS,
        },
    },
    "Standard2019Retail": {
        "product": "Office Standard 2019 (C2R)",
        "supported_versions": ("2019",),
        "architectures": ("x86", "x64"),
        "family": "office",
        "registry_paths": {
            "configuration": C2R_CONFIGURATION_KEYS,
            "product_release_ids": C2R_PRODUCT_RELEASE_ROOTS,
        },
    },
    "ProPlus2021Retail": {
        "product": "Office Professional Plus 2021 (C2R)",
        "supported_versions": ("2021",),
        "architectures": ("x86", "x64", "ARM64"),
        "family": "office",
        "registry_paths": {
            "configuration": C2R_CONFIGURATION_KEYS,
            "product_release_ids": C2R_PRODUCT_RELEASE_ROOTS,
        },
    },
    "Standard2021Retail": {
        "product": "Office Standard 2021 (C2R)",
        "supported_versions": ("2021",),
        "architectures": ("x86", "x64"),
        "family": "office",
        "registry_paths": {
            "configuration": C2R_CONFIGURATION_KEYS,
            "product_release_ids": C2R_PRODUCT_RELEASE_ROOTS,
        },
    },
    "ProPlus2024Retail": {
        "product": "Office Professional Plus 2024 (C2R)",
        "supported_versions": ("2024",),
        "architectures": ("x86", "x64", "ARM64"),
        "family": "office",
        "registry_paths": {
            "configuration": C2R_CONFIGURATION_KEYS,
            "product_release_ids": C2R_PRODUCT_RELEASE_ROOTS,
        },
    },
    "Standard2024Retail": {
        "product": "Office Standard 2024 (C2R)",
        "supported_versions": ("2024",),
        "architectures": ("x86", "x64"),
        "family": "office",
        "registry_paths": {
            "configuration": C2R_CONFIGURATION_KEYS,
            "product_release_ids": C2R_PRODUCT_RELEASE_ROOTS,
        },
    },
    # Project and Visio
    "ProjectProRetail": {
        "product": "Microsoft Project Professional (C2R)",
        "supported_versions": ("2016", "2019", "2021", "2024"),
        "architectures": ("x86", "x64"),
        "family": "project",
        "registry_paths": {
            "configuration": C2R_CONFIGURATION_KEYS,
            "product_release_ids": C2R_PRODUCT_RELEASE_ROOTS,
        },
    },
    "VisioProRetail": {
        "product": "Microsoft Visio Professional (C2R)",
        "supported_versions": ("2016", "2019", "2021", "2024"),
        "architectures": ("x86", "x64"),
        "family": "visio",
        "registry_paths": {
            "configuration": C2R_CONFIGURATION_KEYS,
            "product_release_ids": C2R_PRODUCT_RELEASE_ROOTS,
        },
    },
    # Legacy / hybrid
    "MondoRetail": {
        "product": "Office Mondo (Microsoft Internal)",
        "supported_versions": ("2013", "2016", "2019"),
        "architectures": ("x86", "x64"),
        "family": "office",
        "registry_paths": {
            "configuration": C2R_CONFIGURATION_KEYS,
            "product_release_ids": C2R_PRODUCT_RELEASE_ROOTS,
        },
    },
}


DEFAULT_AUTO_ALL_C2R_RELEASES: Mapping[str, dict[str, object]] = {
    "O365ProPlusRetail": {
        "product": "Microsoft 365 Apps for enterprise",
        "description": "Uninstall Microsoft 365 Apps for enterprise (Click-to-Run)",
        "default_version": "365",
        "channel": "Current Channel",
        "family": "office",
    },
    "O365ProPlusVolume": {
        "product": "Microsoft 365 Apps for enterprise (Volume)",
        "description": "Uninstall Microsoft 365 Apps for enterprise (Volume)",
        "default_version": "365",
        "channel": "Current Channel",
        "family": "office",
    },
    "O365BusinessRetail": {
        "product": "Microsoft 365 Apps for business",
        "description": "Uninstall Microsoft 365 Apps for business",
        "default_version": "365",
        "channel": "Current Channel",
        "family": "office",
    },
    "ProPlus2024Retail": {
        "product": "Office Professional Plus 2024 (C2R)",
        "description": "Uninstall Office Professional Plus 2024 (Click-to-Run)",
        "default_version": "2024",
        "channel": "Perpetual Enterprise",
        "family": "office",
    },
    "Standard2024Retail": {
        "product": "Office Standard 2024 (C2R)",
        "description": "Uninstall Office Standard 2024 (Click-to-Run)",
        "default_version": "2024",
        "channel": "Perpetual Enterprise",
        "family": "office",
    },
    "ProPlus2021Retail": {
        "product": "Office Professional Plus 2021 (C2R)",
        "description": "Uninstall Office Professional Plus 2021 (Click-to-Run)",
        "default_version": "2021",
        "channel": "Perpetual Enterprise",
        "family": "office",
    },
    "Standard2021Retail": {
        "product": "Office Standard 2021 (C2R)",
        "description": "Uninstall Office Standard 2021 (Click-to-Run)",
        "default_version": "2021",
        "channel": "Perpetual Enterprise",
        "family": "office",
    },
    "ProPlus2019Retail": {
        "product": "Office Professional Plus 2019 (C2R)",
        "description": "Uninstall Office Professional Plus 2019 (Click-to-Run)",
        "default_version": "2019",
        "channel": "Perpetual Enterprise",
        "family": "office",
    },
    "Standard2019Retail": {
        "product": "Office Standard 2019 (C2R)",
        "description": "Uninstall Office Standard 2019 (Click-to-Run)",
        "default_version": "2019",
        "channel": "Perpetual Enterprise",
        "family": "office",
    },
    "ProjectProRetail": {
        "product": "Microsoft Project Professional (C2R)",
        "description": "Uninstall Microsoft Project Professional (Click-to-Run)",
        "default_version": "2024",
        "channel": "Perpetual Enterprise",
        "family": "project",
    },
    "VisioProRetail": {
        "product": "Microsoft Visio Professional (C2R)",
        "description": "Uninstall Microsoft Visio Professional (Click-to-Run)",
        "default_version": "2024",
        "channel": "Perpetual Enterprise",
        "family": "visio",
    },
}
"""!
@brief Default Click-to-Run releases seeded for auto-all planning.
@details Provides curated metadata for modern suites that should be targeted
when auto-all mode executes without relying on detection results. Optional
components such as Project and Visio are included so planners can respect
``--include`` selections.
"""


_MSI_FAMILY_LOOKUP: dict[str, str] = {}
for _code, _metadata in MSI_PRODUCT_MAP.items():
    _family = str(_metadata.get("family", ""))
    if not _family:
        continue
    _canonical = _code.upper()
    _MSI_FAMILY_LOOKUP[_canonical] = _family
    _without_braces = _canonical.strip("{}")
    if _without_braces:
        _MSI_FAMILY_LOOKUP[_without_braces] = _family
    _condensed = (
        _without_braces.replace("-", "") if _without_braces else _canonical.replace("-", "")
    )
    if _condensed:
        _MSI_FAMILY_LOOKUP[_condensed] = _family
        _MSI_FAMILY_LOOKUP[f"{{{_condensed}}}"] = _family
"""!
@brief Internal cache mapping MSI product codes to families.
"""

_C2R_FAMILY_LOOKUP = {
    release_id.lower(): str(metadata.get("family", ""))
    for release_id, metadata in C2R_PRODUCT_RELEASES.items()
    if metadata.get("family")
}
"""!
@brief Internal cache mapping Click-to-Run release identifiers to families.
"""

_SUPPORTED_COMPONENT_ALIASES = {
    "msi-project": "project",
    "msi-visio": "visio",
    "c2r-project": "project",
    "c2r-visio": "visio",
    "onenote2016": "onenote",
}
"""!
@brief Normalisation table for optional component identifiers.
"""


def resolve_msi_family(product_code: str | None) -> str | None:
    """!
    @brief Return the product family for the supplied MSI ``product_code``.
    @details Normalises the identifier to ensure braces and case do not affect
    lookups. Returns ``None`` when the product code is unknown.
    """

    if not product_code:
        return None
    normalised = product_code.strip().upper()
    if not normalised:
        return None

    candidates = [normalised]
    stripped = normalised.strip("{}")
    if stripped and stripped not in candidates:
        candidates.append(stripped)
    condensed = stripped.replace("-", "") if stripped else normalised.replace("-", "")
    if condensed and condensed not in candidates:
        candidates.append(condensed)
    if condensed:
        wrapped = f"{{{condensed}}}"
        if wrapped not in candidates:
            candidates.append(wrapped)

    for candidate in candidates:
        family = _MSI_FAMILY_LOOKUP.get(candidate)
        if family:
            return family
    return None


def resolve_c2r_family(release_id: str | None) -> str | None:
    """!
    @brief Return the product family for the supplied Click-to-Run ``release_id``.
    """

    if not release_id:
        return None
    normalised = release_id.strip().lower()
    if not normalised:
        return None
    return _C2R_FAMILY_LOOKUP.get(normalised)


def resolve_supported_component(name: str | None) -> str | None:
    """!
    @brief Normalise a component identifier to a supported component tag.
    @details Accepts entries from :data:`SUPPORTED_COMPONENTS` along with
    historical aliases used by Office removal scripts.
    """

    if not name:
        return None
    candidate = name.strip().lower()
    if not candidate:
        return None
    if candidate in SUPPORTED_COMPONENTS:
        return candidate
    return _SUPPORTED_COMPONENT_ALIASES.get(candidate)


def is_supported_component(name: str | None) -> bool:
    """!
    @brief Convenience predicate returning ``True`` for supported components.
    """

    return resolve_supported_component(name) is not None


def iter_supported_components() -> tuple[str, ...]:
    """!
    @brief Iterate the optional component identifiers recognised by the scrubber.
    """

    return SUPPORTED_COMPONENTS


KNOWN_SCHEDULED_TASKS = (
    r"Microsoft\\Office\\OfficeTelemetryAgentFallBack",
    r"Microsoft\\Office\\OfficeTelemetryAgentLogOn",
    r"Microsoft\\Office\\OfficeBackgroundTaskHandlerLogon",
    r"Microsoft\\Office\\OfficeBackgroundTaskHandlerRegistration",
)

C2R_CLEANUP_TASKS = (
    "FF_INTEGRATEDstreamSchedule",
    "FF_INTEGRATEDUPDATEDETECTION",
    "C2RAppVLoggingStart",
    "Office 15 Subscription Heartbeat",
    "Microsoft Office 15 Sync Maintenance for {d068b555-9700-40b8-992c-f866287b06c1}",
    r"\Microsoft\Office\OfficeInventoryAgentFallBack",
    r"\Microsoft\Office\OfficeTelemetryAgentFallBack",
    r"\Microsoft\Office\OfficeInventoryAgentLogOn",
    r"\Microsoft\Office\OfficeTelemetryAgentLogOn",
    "Office Background Streaming",
    r"\Microsoft\Office\Office Automatic Updates",
    r"\Microsoft\Office\Office ClickToRun Service Monitor",
    "Office Subscription Maintenance",
)
"""!
@brief Scheduled tasks removed by legacy OffScrubC2R flows.
"""

KNOWN_SERVICES = (
    "ClickToRunSvc",
    "OfficeSvc",
    "ose",
    "ose64",
)

LICENSE_DLLS = {
    "spp": "sppc.dll",
    "ospp": "osppc.dll",
}
"""!
@brief Default DLL names referenced by the embedded licensing PowerShell script.
"""

LICENSING_GUID_FILTERS = {
    "office_family": "0ff1ce15-a989-479d-af46-f275c6370663",
}
"""!
@brief Product family GUIDs targeted when uninstalling Office licenses.
"""

OSPP_REGISTRY_PATH = r"HKLM\\SOFTWARE\\Microsoft\\OfficeSoftwareProtectionPlatform"
"""!
@brief Registry location that exposes the Office Software Protection Platform path.
"""

USER_TEMPLATE_PATHS = (
    r"%APPDATA%\\Microsoft\\Templates",
    r"%APPDATA%\\Microsoft\\Office\\Templates",
    r"%LOCALAPPDATA%\\Microsoft\\Office\\Licensing",
    r"%LOCALAPPDATA%\\Microsoft\\Office\\Licenses",
)
"""!
@brief Directories containing user templates and licensing state that require explicit
purge consent.
"""

INSTALL_ROOT_TEMPLATES = (
    # Parent Microsoft Office directories - catch all leftover content after partial uninstall
    {
        "label": "msoffice_root_x86",
        "path": r"C:\\Program Files (x86)\\Microsoft Office",
        "architecture": "x86",
        "release": "all",
    },
    {
        "label": "msoffice_root_x64",
        "path": r"C:\\Program Files\\Microsoft Office",
        "architecture": "x64",
        "release": "all",
    },
    # Specific subdirectories (also listed for granular detection)
    {
        "label": "c2r_root_x86",
        "path": r"C:\\Program Files (x86)\\Microsoft Office\\root",
        "architecture": "x86",
        "release": "C2R",
    },
    {
        "label": "c2r_root_x64",
        "path": r"C:\\Program Files\\Microsoft Office\\root",
        "architecture": "x64",
        "release": "C2R",
    },
    {
        "label": "office16_x86",
        "path": r"C:\\Program Files (x86)\\Microsoft Office\\Office16",
        "architecture": "x86",
        "release": "2016",
    },
    {
        "label": "office16_x64",
        "path": r"C:\\Program Files\\Microsoft Office\\Office16",
        "architecture": "x64",
        "release": "2016",
    },
    {
        "label": "office15_x86",
        "path": r"C:\\Program Files (x86)\\Microsoft Office\\Office15",
        "architecture": "x86",
        "release": "2013",
    },
    {
        "label": "office15_x64",
        "path": r"C:\\Program Files\\Microsoft Office\\Office15",
        "architecture": "x64",
        "release": "2013",
    },
    {
        "label": "office14_x86",
        "path": r"C:\\Program Files (x86)\\Microsoft Office\\Office14",
        "architecture": "x86",
        "release": "2010",
    },
    {
        "label": "office14_x64",
        "path": r"C:\\Program Files\\Microsoft Office\\Office14",
        "architecture": "x64",
        "release": "2010",
    },
    {
        "label": "office12_x86",
        "path": r"C:\\Program Files (x86)\\Microsoft Office\\Office12",
        "architecture": "x86",
        "release": "2007",
    },
    {
        "label": "office12_x64",
        "path": r"C:\\Program Files\\Microsoft Office\\Office12",
        "architecture": "x64",
        "release": "2007",
    },
    # Common Files shared Office components
    {
        "label": "shared_office_x86",
        "path": r"C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\OFFICE16",
        "architecture": "x86",
        "release": "2016",
    },
    {
        "label": "shared_office_x64",
        "path": r"C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE16",
        "architecture": "x64",
        "release": "2016",
    },
    {
        "label": "shared_office15_x86",
        "path": r"C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\OFFICE15",
        "architecture": "x86",
        "release": "2013",
    },
    {
        "label": "shared_office15_x64",
        "path": r"C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE15",
        "architecture": "x64",
        "release": "2013",
    },
    {
        "label": "shared_office14_x86",
        "path": r"C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\OFFICE14",
        "architecture": "x86",
        "release": "2010",
    },
    {
        "label": "shared_office14_x64",
        "path": r"C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE14",
        "architecture": "x64",
        "release": "2010",
    },
    {
        "label": "shared_office12_x86",
        "path": r"C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\OFFICE12",
        "architecture": "x86",
        "release": "2007",
    },
    {
        "label": "shared_office12_x64",
        "path": r"C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE12",
        "architecture": "x64",
        "release": "2007",
    },
)

__all__ = [
    "ALL_OFFICE_PROCESSES",
    "C2R_CHANNEL_ALIASES",
    "C2R_CLEANUP_TASKS",
    "C2R_COM_REGISTRY_PATHS",
    "C2R_CONFIGURATION_KEYS",
    "C2R_INFRASTRUCTURE_PROCESSES",
    "C2R_OFFSCRUB_ARGS",
    "C2R_OFFSCRUB_SCRIPT",
    "C2R_PLATFORM_ALIASES",
    "C2R_PRODUCT_RELEASES",
    "DEFAULT_AUTO_ALL_C2R_RELEASES",
    "C2R_PRODUCT_RELEASE_ROOTS",
    "C2R_SUBSCRIPTION_ROOTS",
    "C2R_UNINSTALL_VERSION_GROUPS",
    "DEFAULT_OFFICE_PROCESSES",
    "OFFICE_PROCESS_PATTERNS",
    "HKCR",
    "HKCU",
    "HKLM",
    "HKU",
    "INSTALL_ROOT_TEMPLATES",
    "KNOWN_SCHEDULED_TASKS",
    "KNOWN_SERVICES",
    "LICENSE_DLLS",
    "LICENSING_GUID_FILTERS",
    "MSI_INFRASTRUCTURE_PROCESSES",
    "MSI_UNINSTALL_VERSION_GROUPS",
    "MSI_OFFSCRUB_ARGS",
    "MSI_OFFSCRUB_DEFAULT_SCRIPT",
    "MSI_OFFSCRUB_SCRIPT_MAP",
    "MSI_PRODUCT_MAP",
    "MSI_UNINSTALL_ROOTS",
    "MSIEXEC_RETURN_CODES",
    "OFFSCRUB_UNINSTALL_PRIORITY",
    "OFFSCRUB_UNINSTALL_SEQUENCE",
    "OFFICE_SCHEDULED_TASKS_TO_DELETE",
    "OFFICE_SERVICES_TO_DELETE",
    "OSPP_REGISTRY_PATH",
    "OFFSCRUB_EXECUTABLE",
    "OFFSCRUB_HOST_ARGS",
    "REGISTRY_RESIDUE_PATHS",
    "REGISTRY_ROOTS",
    "RESIDUE_PATH_TEMPLATES",
    "SUPPORTED_COMPONENTS",
    "SUPPORTED_TARGETS",
    "SUPPORTED_VERSIONS",
    "UNINSTALL_COMMAND_TEMPLATES",
    "USER_TEMPLATE_PATHS",
    "known_msi_codes",
    "resolve_c2r_family",
    "resolve_msi_family",
    "resolve_supported_component",
    "is_supported_component",
    "iter_supported_components",
    "translate_msiexec_return_code",
]
