"""!
@brief Static data and enumerations for Office Janitor.
@details Centralises product identifiers, registry roots, Click-to-Run channel
metadata, and other shared constants so detection and uninstall modules work
from a single, versioned source of truth.
"""
from __future__ import annotations

from typing import Dict, Iterable, Mapping, Tuple

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


def _merge_roots(*roots: Tuple[int, str]) -> Tuple[Tuple[int, str], ...]:
    """!
    @brief Provide a helper that returns a deduplicated tuple of uninstall roots.
    """

    seen: set[Tuple[int, str]] = set()
    ordered: list[Tuple[int, str]] = []
    for hive, path in roots:
        entry = (hive, path)
        if entry in seen:
            continue
        ordered.append(entry)
        seen.add(entry)
    return tuple(ordered)


MSI_PRODUCT_MAP: Dict[str, Dict[str, object]] = {
    # Office 2010 (14.x) Professional Plus
    "{90140000-0011-0000-0000-0000000FF1CE}": {
        "product": "Microsoft Office Professional Plus 2010",
        "edition": "Professional Plus",
        "version": "2010",
        "supported_versions": ("2010",),
        "architecture": "x86",
        "registry_roots": MSI_UNINSTALL_ROOTS,
    },
    "{90140000-0011-0000-1000-0000000FF1CE}": {
        "product": "Microsoft Office Professional Plus 2010",
        "edition": "Professional Plus",
        "version": "2010",
        "supported_versions": ("2010",),
        "architecture": "x64",
        "registry_roots": _merge_roots(MSI_UNINSTALL_ROOTS[0]),
    },
    # Office 2013 (15.x) Professional Plus
    "{90150000-0011-0000-0000-0000000FF1CE}": {
        "product": "Microsoft Office Professional Plus 2013",
        "edition": "Professional Plus",
        "version": "2013",
        "supported_versions": ("2013",),
        "architecture": "x86",
        "registry_roots": MSI_UNINSTALL_ROOTS,
    },
    "{90150000-0011-0000-1000-0000000FF1CE}": {
        "product": "Microsoft Office Professional Plus 2013",
        "edition": "Professional Plus",
        "version": "2013",
        "supported_versions": ("2013",),
        "architecture": "x64",
        "registry_roots": _merge_roots(MSI_UNINSTALL_ROOTS[0]),
    },
    # Office 2016/2019/2021/2024 perpetual channel (MSI-based SKUs)
    "{90160000-0011-0000-0000-0000000FF1CE}": {
        "product": "Microsoft Office Professional Plus 2016",
        "edition": "Professional Plus",
        "version": "2016",
        "supported_versions": ("2016", "2019", "2021", "2024"),
        "architecture": "x86",
        "registry_roots": MSI_UNINSTALL_ROOTS,
    },
    "{90160000-0011-0000-1000-0000000FF1CE}": {
        "product": "Microsoft Office Professional Plus 2016",
        "edition": "Professional Plus",
        "version": "2016",
        "supported_versions": ("2016", "2019", "2021", "2024"),
        "architecture": "x64",
        "registry_roots": _merge_roots(MSI_UNINSTALL_ROOTS[0]),
    },
    # Visio Professional 2016 (MSI)
    "{90160000-0051-0000-0000-0000000FF1CE}": {
        "product": "Microsoft Visio Professional 2016",
        "edition": "Visio Professional",
        "version": "2016",
        "supported_versions": ("2016", "2019", "2021", "2024"),
        "architecture": "x86",
        "registry_roots": MSI_UNINSTALL_ROOTS,
    },
    "{90160000-0051-0000-1000-0000000FF1CE}": {
        "product": "Microsoft Visio Professional 2016",
        "edition": "Visio Professional",
        "version": "2016",
        "supported_versions": ("2016", "2019", "2021", "2024"),
        "architecture": "x64",
        "registry_roots": _merge_roots(MSI_UNINSTALL_ROOTS[0]),
    },
    # Project Professional 2016 (MSI)
    "{90160000-003B-0000-0000-0000000FF1CE}": {
        "product": "Microsoft Project Professional 2016",
        "edition": "Project Professional",
        "version": "2016",
        "supported_versions": ("2016", "2019", "2021", "2024"),
        "architecture": "x86",
        "registry_roots": MSI_UNINSTALL_ROOTS,
    },
    "{90160000-003B-0000-1000-0000000FF1CE}": {
        "product": "Microsoft Project Professional 2016",
        "edition": "Project Professional",
        "version": "2016",
        "supported_versions": ("2016", "2019", "2021", "2024"),
        "architecture": "x64",
        "registry_roots": _merge_roots(MSI_UNINSTALL_ROOTS[0]),
    },
}


def known_msi_codes() -> Iterable[str]:
    """!
    @brief Iterate the product codes that map to Office MSI deployments.
    """

    return MSI_PRODUCT_MAP.keys()


C2R_CONFIGURATION_KEYS: Tuple[Tuple[int, str], ...] = (
    (HKLM, r"SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration"),
    (HKLM, r"SOFTWARE\\Microsoft\\Office\\15.0\\ClickToRun\\Configuration"),
    (HKLM, r"SOFTWARE\\WOW6432Node\\Microsoft\\Office\\15.0\\ClickToRun\\Configuration"),
    (HKCU, r"SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration"),
)

C2R_PRODUCT_RELEASE_ROOTS: Tuple[Tuple[int, str], ...] = (
    (HKLM, r"SOFTWARE\\Microsoft\\Office\\ClickToRun\\ProductReleaseIDs"),
    (HKLM, r"SOFTWARE\\Microsoft\\Office\\15.0\\ClickToRun\\ProductReleaseIDs"),
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
    "amd64": "x64",
    "arm64": "ARM64",
    "neutral": "neutral",
}

C2R_CHANNEL_ALIASES: Dict[str, str] = {
    "Production::CC": "Current Channel",
    "Production::Current": "Current Channel",
    "Production::MEC": "Monthly Enterprise Channel",
    "Production::SAEC": "Semi-Annual Enterprise Channel",
    "Production::SA": "Semi-Annual Channel",
    "Production::Beta": "Beta Channel",
    "Production::InsiderFast": "Insider Fast",
    "Production::FirstReleaseCurrent": "Current Channel (Preview)",
    "Production::FirstReleaseDeferred": "Semi-Annual Preview",
    "http://officecdn.microsoft.com/pr/492350f6-3a04-4b59-8b34-4c547755c2a0": "Current Channel",
    "http://officecdn.microsoft.com/pr/55336b82-a18d-4dd6-b5f6-9e5095c314a6": "Monthly Enterprise Channel",
    "http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114": "Semi-Annual Enterprise Channel",
    "http://officecdn.microsoft.com/pr/5440fd1f-7ecb-4221-8110-14e4edeeb5d0": "Semi-Annual Preview",
}

C2R_PRODUCT_RELEASES: Mapping[str, Dict[str, object]] = {
    # Microsoft 365 Apps / Office 2016+ suites
    "O365ProPlusRetail": {
        "product": "Microsoft 365 Apps for enterprise",
        "supported_versions": ("2016", "2019", "2021", "2024", "365"),
        "architectures": ("x86", "x64", "ARM64"),
        "registry_paths": {
            "configuration": C2R_CONFIGURATION_KEYS,
            "product_release_ids": C2R_PRODUCT_RELEASE_ROOTS,
        },
    },
    "O365ProPlusVolume": {
        "product": "Microsoft 365 Apps for enterprise (Volume)",
        "supported_versions": ("2016", "2019", "2021", "2024", "365"),
        "architectures": ("x86", "x64", "ARM64"),
        "registry_paths": {
            "configuration": C2R_CONFIGURATION_KEYS,
            "product_release_ids": C2R_PRODUCT_RELEASE_ROOTS,
        },
    },
    "O365BusinessRetail": {
        "product": "Microsoft 365 Apps for business",
        "supported_versions": ("2016", "2019", "2021", "2024", "365"),
        "architectures": ("x86", "x64", "ARM64"),
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
        "registry_paths": {
            "configuration": C2R_CONFIGURATION_KEYS,
            "product_release_ids": C2R_PRODUCT_RELEASE_ROOTS,
        },
    },
    "Standard2019Retail": {
        "product": "Office Standard 2019 (C2R)",
        "supported_versions": ("2019",),
        "architectures": ("x86", "x64"),
        "registry_paths": {
            "configuration": C2R_CONFIGURATION_KEYS,
            "product_release_ids": C2R_PRODUCT_RELEASE_ROOTS,
        },
    },
    "ProPlus2021Retail": {
        "product": "Office Professional Plus 2021 (C2R)",
        "supported_versions": ("2021",),
        "architectures": ("x86", "x64", "ARM64"),
        "registry_paths": {
            "configuration": C2R_CONFIGURATION_KEYS,
            "product_release_ids": C2R_PRODUCT_RELEASE_ROOTS,
        },
    },
    "Standard2021Retail": {
        "product": "Office Standard 2021 (C2R)",
        "supported_versions": ("2021",),
        "architectures": ("x86", "x64"),
        "registry_paths": {
            "configuration": C2R_CONFIGURATION_KEYS,
            "product_release_ids": C2R_PRODUCT_RELEASE_ROOTS,
        },
    },
    "ProPlus2024Retail": {
        "product": "Office Professional Plus 2024 (C2R)",
        "supported_versions": ("2024",),
        "architectures": ("x86", "x64", "ARM64"),
        "registry_paths": {
            "configuration": C2R_CONFIGURATION_KEYS,
            "product_release_ids": C2R_PRODUCT_RELEASE_ROOTS,
        },
    },
    "Standard2024Retail": {
        "product": "Office Standard 2024 (C2R)",
        "supported_versions": ("2024",),
        "architectures": ("x86", "x64"),
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
        "registry_paths": {
            "configuration": C2R_CONFIGURATION_KEYS,
            "product_release_ids": C2R_PRODUCT_RELEASE_ROOTS,
        },
    },
    "VisioProRetail": {
        "product": "Microsoft Visio Professional (C2R)",
        "supported_versions": ("2016", "2019", "2021", "2024"),
        "architectures": ("x86", "x64"),
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
        "registry_paths": {
            "configuration": C2R_CONFIGURATION_KEYS,
            "product_release_ids": C2R_PRODUCT_RELEASE_ROOTS,
        },
    },
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
    "C2R_CHANNEL_ALIASES",
    "C2R_COM_REGISTRY_PATHS",
    "C2R_CONFIGURATION_KEYS",
    "C2R_PLATFORM_ALIASES",
    "C2R_PRODUCT_RELEASES",
    "C2R_PRODUCT_RELEASE_ROOTS",
    "C2R_SUBSCRIPTION_ROOTS",
    "DEFAULT_OFFICE_PROCESSES",
    "HKCR",
    "HKCU",
    "HKLM",
    "HKU",
    "INSTALL_ROOT_TEMPLATES",
    "KNOWN_SCHEDULED_TASKS",
    "KNOWN_SERVICES",
    "MSI_PRODUCT_MAP",
    "MSI_UNINSTALL_ROOTS",
    "REGISTRY_ROOTS",
    "SUPPORTED_VERSIONS",
    "known_msi_codes",
]
