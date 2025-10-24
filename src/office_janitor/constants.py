"""!
@brief Static data and enumerations for Office Janitor.
@details Holds product code mappings, registry roots, default paths, and other
constants used across detection, planning, and scrub orchestration per the
specification.
"""
from __future__ import annotations

from typing import Dict, Tuple

try:  # pragma: no cover - Windows registry handles are optional on test hosts.
    import winreg
except ImportError:  # pragma: no cover - test scaffolding supplies substitutes.
    winreg = None  # type: ignore[assignment]


if winreg is not None:  # pragma: no branch - deterministic assignments.
    HKLM = winreg.HKEY_LOCAL_MACHINE
    HKCU = winreg.HKEY_CURRENT_USER
    HKCR = winreg.HKEY_CLASSES_ROOT
    HKU = winreg.HKEY_USERS
else:  # pragma: no cover - exercised implicitly in non-Windows CI.
    HKLM = 0x80000002
    HKCU = 0x80000001
    HKCR = 0x80000000
    HKU = 0x80000003


REGISTRY_ROOTS: Dict[str, int] = {
    "HKLM": HKLM,
    "HKCU": HKCU,
    "HKCR": HKCR,
    "HKU": HKU,
}

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

DEFAULT_OFFICE_PROCESSES = (
    "winword.exe",
    "excel.exe",
    "outlook.exe",
    "onenote.exe",
    "visio.exe",
    "powerpnt.exe",
)

MSI_UNINSTALL_ROOTS: Tuple[Tuple[int, str], ...] = (
    (HKLM, r"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall"),
    (HKLM, r"SOFTWARE\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall"),
)

MSI_PRODUCT_CODES: Dict[str, Dict[str, str]] = {
    "{90150000-0011-0000-0000-0000000FF1CE}": {
        "release": "2013",
        "generation": "Office15",
        "edition": "Professional Plus",
        "architecture": "x86",
        "install_path": r"C:\\Program Files (x86)\\Microsoft Office\\Office15",
    },
    "{91160000-0011-0000-0000-0000000FF1CE}": {
        "release": "2016",
        "generation": "Office16",
        "edition": "Professional Plus",
        "architecture": "x64",
        "install_path": r"C:\\Program Files\\Microsoft Office\\Office16",
    },
    "{91190000-0011-0000-0000-0000000FF1CE}": {
        "release": "2019",
        "generation": "Office17",
        "edition": "Professional Plus",
        "architecture": "x64",
        "install_path": r"C:\\Program Files\\Microsoft Office\\Office17",
    },
    "{91140000-0011-0000-0000-0000000FF1CE}": {
        "release": "2010",
        "generation": "Office14",
        "edition": "Professional Plus",
        "architecture": "x86",
        "install_path": r"C:\\Program Files (x86)\\Microsoft Office\\Office14",
    },
}

C2R_CONFIGURATION_KEYS: Tuple[Tuple[int, str], ...] = (
    (HKLM, r"SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration"),
    (HKCU, r"SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration"),
)

C2R_SUBSCRIPTION_ROOTS: Tuple[Tuple[int, str], ...] = (
    (HKLM, r"SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration\\Subscriptions"),
    (HKCU, r"SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration\\Subscriptions"),
)

C2R_COM_REGISTRY_PATHS: Tuple[Tuple[int, str], ...] = (
    (HKLM, r"SOFTWARE\\Microsoft\\Office\\ClickToRun\\COM Compatibility\\Applications"),
    (HKCU, r"SOFTWARE\\Microsoft\\Office\\ClickToRun\\COM Compatibility\\Applications"),
)

C2R_PLATFORM_ALIASES: Dict[str, str] = {
    "x86": "x86",
    "x64": "x64",
    "arm64": "ARM64",
    "neutral": "neutral",
}

C2R_CHANNELS: Dict[str, str] = {
    "Production::CC": "Current Channel",
    "Production::MEC": "Monthly Enterprise Channel",
    "Production::SAEC": "Semi-Annual Enterprise Channel",
    "Production::Beta": "Beta Channel",
    "Production::InsiderFast": "Insider Fast",
    "http://officecdn.microsoft.com/pr/492350f6-3a04-4b59-8b34-4c547755c2a0": "Current Channel",
    "http://officecdn.microsoft.com/pr/55336b82-a18d-4dd6-b5f6-9e5095c314a6": "Monthly Enterprise Channel",
    "http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114": "Semi-Annual Enterprise Channel",
}

KNOWN_SCHEDULED_TASKS = (
    r"Microsoft\\Office\\OfficeTelemetryAgentFallBack",
    r"Microsoft\\Office\\OfficeTelemetryAgentLogOn",
    r"Microsoft\\Office\\OfficeBackgroundTaskHandlerLogon",
    r"Microsoft\\Office\\OfficeBackgroundTaskHandlerRegistration",
)

KNOWN_SERVICES = (
    "ClickToRunSvc",
    "OfficeSvc",
    "ose",
    "ose64",
)

INSTALL_ROOT_TEMPLATES = (
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
)

__all__ = [
    "C2R_CHANNELS",
    "C2R_COM_REGISTRY_PATHS",
    "C2R_CONFIGURATION_KEYS",
    "C2R_PLATFORM_ALIASES",
    "C2R_SUBSCRIPTION_ROOTS",
    "DEFAULT_OFFICE_PROCESSES",
    "HKCR",
    "HKCU",
    "HKLM",
    "HKU",
    "INSTALL_ROOT_TEMPLATES",
    "KNOWN_SCHEDULED_TASKS",
    "KNOWN_SERVICES",
    "MSI_PRODUCT_CODES",
    "MSI_UNINSTALL_ROOTS",
    "REGISTRY_ROOTS",
    "SUPPORTED_VERSIONS",
]
