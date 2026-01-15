"""!
@brief Office Deployment Tool XML configuration builder.
@details Generates ODT XML configuration files for installing Microsoft Office
products. Supports all product IDs and update channels available in the official
OEM configurations, allowing users to create custom installation configurations
programmatically.

@see https://docs.microsoft.com/en-us/deployoffice/office-deployment-tool-configuration-options
"""

from __future__ import annotations

import glob
import os
import re
import subprocess
import sys
import tempfile
import threading
import time
import xml.etree.ElementTree as ET
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from types import ModuleType

from . import logging_ext

# Try to import spinner for progress display
_spinner: ModuleType | None
try:
    from . import spinner as _spinner
except ImportError:
    _spinner = None

# ---------------------------------------------------------------------------
# Constants and Enumerations
# ---------------------------------------------------------------------------


class UpdateChannel(Enum):
    """!
    @brief Office update channels supported by ODT.
    @details Maps friendly names to ODT channel identifiers.
    @see https://docs.microsoft.com/en-us/deployoffice/overview-update-channels
    """

    # Microsoft 365 Apps channels
    CURRENT = "Current"
    CURRENT_PREVIEW = "CurrentPreview"
    MONTHLY_ENTERPRISE = "MonthlyEnterprise"
    SEMI_ANNUAL_PREVIEW = "SemiAnnualPreview"
    SEMI_ANNUAL = "SemiAnnual"
    BETA = "BetaChannel"

    # Perpetual (LTSC) channels
    PERPETUAL_VL_2024 = "PerpetualVL2024"
    PERPETUAL_VL_2021 = "PerpetualVL2021"
    PERPETUAL_VL_2019 = "PerpetualVL2019"

    # Legacy
    INSIDER_FAST = "InsiderFast"


class Architecture(Enum):
    """!
    @brief Supported Office architectures.
    """

    X86 = "32"
    X64 = "64"


class DisplayLevel(Enum):
    """!
    @brief Installation UI visibility options.
    """

    NONE = "None"
    FULL = "Full"


# Product ID definitions matching official ODT documentation
PRODUCT_IDS: Mapping[str, dict[str, object]] = {
    # Microsoft 365 Apps subscriptions
    "O365ProPlusRetail": {
        "name": "Microsoft 365 Apps for enterprise",
        "channels": [
            UpdateChannel.CURRENT,
            UpdateChannel.MONTHLY_ENTERPRISE,
            UpdateChannel.SEMI_ANNUAL,
        ],
        "description": "Full-featured Office suite for enterprise (subscription)",
    },
    "O365BusinessRetail": {
        "name": "Microsoft 365 Apps for business",
        "channels": [UpdateChannel.CURRENT, UpdateChannel.MONTHLY_ENTERPRISE],
        "description": "Office suite for small/medium business (subscription)",
    },
    "O365HomePremRetail": {
        "name": "Microsoft 365 Family/Personal",
        "channels": [UpdateChannel.CURRENT],
        "description": "Consumer Microsoft 365 subscription",
    },
    "O365SmallBusPremRetail": {
        "name": "Microsoft 365 Business Basic/Standard",
        "channels": [UpdateChannel.CURRENT],
        "description": "Microsoft 365 Business plans",
    },
    "O365EduCloudRetail": {
        "name": "Microsoft 365 Education",
        "channels": [UpdateChannel.CURRENT, UpdateChannel.SEMI_ANNUAL],
        "description": "Microsoft 365 for Education",
    },
    # Office 2024 LTSC (Volume License)
    "ProPlus2024Volume": {
        "name": "Office LTSC Professional Plus 2024",
        "channels": [UpdateChannel.PERPETUAL_VL_2024],
        "description": "Perpetual Office 2024 for enterprise (volume)",
    },
    "Standard2024Volume": {
        "name": "Office LTSC Standard 2024",
        "channels": [UpdateChannel.PERPETUAL_VL_2024],
        "description": "Perpetual Office 2024 Standard (volume)",
    },
    "ProPlus2024Retail": {
        "name": "Office Professional Plus 2024",
        "channels": [UpdateChannel.PERPETUAL_VL_2024],
        "description": "Perpetual Office 2024 for enterprise (retail)",
    },
    "Standard2024Retail": {
        "name": "Office Standard 2024",
        "channels": [UpdateChannel.PERPETUAL_VL_2024],
        "description": "Perpetual Office 2024 Standard (retail)",
    },
    # Office 2021 LTSC
    "ProPlus2021Volume": {
        "name": "Office LTSC Professional Plus 2021",
        "channels": [UpdateChannel.PERPETUAL_VL_2021],
        "description": "Perpetual Office 2021 for enterprise (volume)",
    },
    "Standard2021Volume": {
        "name": "Office LTSC Standard 2021",
        "channels": [UpdateChannel.PERPETUAL_VL_2021],
        "description": "Perpetual Office 2021 Standard (volume)",
    },
    "ProPlus2021Retail": {
        "name": "Office Professional Plus 2021",
        "channels": [UpdateChannel.PERPETUAL_VL_2021],
        "description": "Perpetual Office 2021 for enterprise (retail)",
    },
    "Standard2021Retail": {
        "name": "Office Standard 2021",
        "channels": [UpdateChannel.PERPETUAL_VL_2021],
        "description": "Perpetual Office 2021 Standard (retail)",
    },
    # Office 2019
    "ProPlus2019Volume": {
        "name": "Office Professional Plus 2019",
        "channels": [UpdateChannel.PERPETUAL_VL_2019],
        "description": "Perpetual Office 2019 for enterprise (volume)",
    },
    "Standard2019Volume": {
        "name": "Office Standard 2019",
        "channels": [UpdateChannel.PERPETUAL_VL_2019],
        "description": "Perpetual Office 2019 Standard (volume)",
    },
    "ProPlus2019Retail": {
        "name": "Office Professional Plus 2019",
        "channels": [UpdateChannel.PERPETUAL_VL_2019],
        "description": "Perpetual Office 2019 (retail)",
    },
    "Standard2019Retail": {
        "name": "Office Standard 2019",
        "channels": [UpdateChannel.PERPETUAL_VL_2019],
        "description": "Perpetual Office 2019 Standard (retail)",
    },
    # Project products
    "ProjectProRetail": {
        "name": "Project Professional",
        "channels": [
            UpdateChannel.CURRENT,
            UpdateChannel.PERPETUAL_VL_2024,
            UpdateChannel.PERPETUAL_VL_2021,
        ],
        "description": "Microsoft Project Professional (subscription/retail)",
    },
    "ProjectStdRetail": {
        "name": "Project Standard",
        "channels": [
            UpdateChannel.CURRENT,
            UpdateChannel.PERPETUAL_VL_2024,
            UpdateChannel.PERPETUAL_VL_2021,
        ],
        "description": "Microsoft Project Standard",
    },
    "ProjectPro2024Volume": {
        "name": "Project Professional 2024",
        "channels": [UpdateChannel.PERPETUAL_VL_2024],
        "description": "Microsoft Project 2024 Professional (volume)",
    },
    "ProjectStd2024Volume": {
        "name": "Project Standard 2024",
        "channels": [UpdateChannel.PERPETUAL_VL_2024],
        "description": "Microsoft Project 2024 Standard (volume)",
    },
    "ProjectPro2024Retail": {
        "name": "Project Professional 2024",
        "channels": [UpdateChannel.PERPETUAL_VL_2024],
        "description": "Microsoft Project 2024 Professional (retail)",
    },
    "ProjectPro2021Volume": {
        "name": "Project Professional 2021",
        "channels": [UpdateChannel.PERPETUAL_VL_2021],
        "description": "Microsoft Project 2021 Professional (volume)",
    },
    "ProjectStd2021Volume": {
        "name": "Project Standard 2021",
        "channels": [UpdateChannel.PERPETUAL_VL_2021],
        "description": "Microsoft Project 2021 Standard (volume)",
    },
    "ProjectPro2021Retail": {
        "name": "Project Professional 2021",
        "channels": [UpdateChannel.PERPETUAL_VL_2021],
        "description": "Microsoft Project 2021 Professional (retail)",
    },
    "ProjectPro2019Volume": {
        "name": "Project Professional 2019",
        "channels": [UpdateChannel.PERPETUAL_VL_2019],
        "description": "Microsoft Project 2019 Professional (volume)",
    },
    "ProjectPro2019Retail": {
        "name": "Project Professional 2019",
        "channels": [UpdateChannel.PERPETUAL_VL_2019],
        "description": "Microsoft Project 2019 Professional (retail)",
    },
    # Visio products
    "VisioProRetail": {
        "name": "Visio Professional",
        "channels": [
            UpdateChannel.CURRENT,
            UpdateChannel.PERPETUAL_VL_2024,
            UpdateChannel.PERPETUAL_VL_2021,
        ],
        "description": "Microsoft Visio Professional (subscription/retail)",
    },
    "VisioStdRetail": {
        "name": "Visio Standard",
        "channels": [
            UpdateChannel.CURRENT,
            UpdateChannel.PERPETUAL_VL_2024,
            UpdateChannel.PERPETUAL_VL_2021,
        ],
        "description": "Microsoft Visio Standard",
    },
    "VisioPro2024Volume": {
        "name": "Visio Professional 2024",
        "channels": [UpdateChannel.PERPETUAL_VL_2024],
        "description": "Microsoft Visio 2024 Professional (volume)",
    },
    "VisioStd2024Volume": {
        "name": "Visio Standard 2024",
        "channels": [UpdateChannel.PERPETUAL_VL_2024],
        "description": "Microsoft Visio 2024 Standard (volume)",
    },
    "VisioPro2024Retail": {
        "name": "Visio Professional 2024",
        "channels": [UpdateChannel.PERPETUAL_VL_2024],
        "description": "Microsoft Visio 2024 Professional (retail)",
    },
    "VisioPro2021Volume": {
        "name": "Visio Professional 2021",
        "channels": [UpdateChannel.PERPETUAL_VL_2021],
        "description": "Microsoft Visio 2021 Professional (volume)",
    },
    "VisioStd2021Volume": {
        "name": "Visio Standard 2021",
        "channels": [UpdateChannel.PERPETUAL_VL_2021],
        "description": "Microsoft Visio 2021 Standard (volume)",
    },
    "VisioPro2021Retail": {
        "name": "Visio Professional 2021",
        "channels": [UpdateChannel.PERPETUAL_VL_2021],
        "description": "Microsoft Visio 2021 Professional (retail)",
    },
    "VisioPro2019Volume": {
        "name": "Visio Professional 2019",
        "channels": [UpdateChannel.PERPETUAL_VL_2019],
        "description": "Microsoft Visio 2019 Professional (volume)",
    },
    "VisioPro2019Retail": {
        "name": "Visio Professional 2019",
        "channels": [UpdateChannel.PERPETUAL_VL_2019],
        "description": "Microsoft Visio 2019 Professional (retail)",
    },
    # Access Runtime
    "AccessRuntimeRetail": {
        "name": "Access Runtime",
        "channels": [
            UpdateChannel.CURRENT,
            UpdateChannel.PERPETUAL_VL_2024,
            UpdateChannel.PERPETUAL_VL_2021,
        ],
        "description": "Microsoft Access Runtime (for database distribution)",
    },
    # Language packs
    "LanguagePack": {
        "name": "Office Language Pack",
        "channels": [
            UpdateChannel.CURRENT,
            UpdateChannel.PERPETUAL_VL_2024,
            UpdateChannel.PERPETUAL_VL_2021,
        ],
        "description": "Additional language pack for Office",
    },
    # Proofing tools
    "ProofingTools": {
        "name": "Office Proofing Tools",
        "channels": [
            UpdateChannel.CURRENT,
            UpdateChannel.PERPETUAL_VL_2024,
            UpdateChannel.PERPETUAL_VL_2021,
        ],
        "description": "Spelling and grammar tools for additional languages",
    },
}
"""!
@brief Complete product ID catalog matching ODT documentation.
"""

# Supported language/culture codes
SUPPORTED_LANGUAGES: tuple[str, ...] = (
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
    "fr-fr",  # French (France)
    "fr-ca",  # French (Canada)
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
    "es-mx",  # Spanish (Mexico)
    "sv-se",  # Swedish
    "th-th",  # Thai
    "tr-tr",  # Turkish
    "uk-ua",  # Ukrainian
    "vi-vn",  # Vietnamese
)
"""!
@brief Language codes supported for Office installation.
"""

# Pre-built configuration presets
INSTALL_PRESETS: Mapping[str, dict[str, object]] = {
    "365-proplus-x64": {
        "products": ["O365ProPlusRetail"],
        "architecture": Architecture.X64,
        "channel": UpdateChannel.CURRENT,
        "description": "Microsoft 365 Apps for enterprise (64-bit)",
    },
    "365-proplus-x86": {
        "products": ["O365ProPlusRetail"],
        "architecture": Architecture.X86,
        "channel": UpdateChannel.CURRENT,
        "description": "Microsoft 365 Apps for enterprise (32-bit)",
    },
    "365-business-x64": {
        "products": ["O365BusinessRetail"],
        "architecture": Architecture.X64,
        "channel": UpdateChannel.CURRENT,
        "description": "Microsoft 365 Apps for business (64-bit)",
    },
    "365-proplus-visio-project": {
        "products": ["O365ProPlusRetail", "VisioProRetail", "ProjectProRetail"],
        "architecture": Architecture.X64,
        "channel": UpdateChannel.CURRENT,
        "description": "Microsoft 365 Apps + Visio + Project (64-bit)",
    },
    "office2024-x64": {
        "products": ["ProPlus2024Volume"],
        "architecture": Architecture.X64,
        "channel": UpdateChannel.PERPETUAL_VL_2024,
        "description": "Office LTSC 2024 Professional Plus (64-bit)",
    },
    "office2024-x86": {
        "products": ["ProPlus2024Volume"],
        "architecture": Architecture.X86,
        "channel": UpdateChannel.PERPETUAL_VL_2024,
        "description": "Office LTSC 2024 Professional Plus (32-bit)",
    },
    "office2024-standard-x64": {
        "products": ["Standard2024Volume"],
        "architecture": Architecture.X64,
        "channel": UpdateChannel.PERPETUAL_VL_2024,
        "description": "Office LTSC 2024 Standard (64-bit)",
    },
    "office2021-x64": {
        "products": ["ProPlus2021Volume"],
        "architecture": Architecture.X64,
        "channel": UpdateChannel.PERPETUAL_VL_2021,
        "description": "Office LTSC 2021 Professional Plus (64-bit)",
    },
    "office2021-x86": {
        "products": ["ProPlus2021Volume"],
        "architecture": Architecture.X86,
        "channel": UpdateChannel.PERPETUAL_VL_2021,
        "description": "Office LTSC 2021 Professional Plus (32-bit)",
    },
    "office2021-standard-x64": {
        "products": ["Standard2021Volume"],
        "architecture": Architecture.X64,
        "channel": UpdateChannel.PERPETUAL_VL_2021,
        "description": "Office LTSC 2021 Standard (64-bit)",
    },
    "office2019-x64": {
        "products": ["ProPlus2019Volume"],
        "architecture": Architecture.X64,
        "channel": UpdateChannel.PERPETUAL_VL_2019,
        "description": "Office 2019 Professional Plus (64-bit)",
    },
    "office2019-x86": {
        "products": ["ProPlus2019Volume"],
        "architecture": Architecture.X86,
        "channel": UpdateChannel.PERPETUAL_VL_2019,
        "description": "Office 2019 Professional Plus (32-bit)",
    },
    "visio-pro-x64": {
        "products": ["VisioPro2024Volume"],
        "architecture": Architecture.X64,
        "channel": UpdateChannel.PERPETUAL_VL_2024,
        "description": "Visio Professional 2024 (64-bit)",
    },
    "project-pro-x64": {
        "products": ["ProjectPro2024Volume"],
        "architecture": Architecture.X64,
        "channel": UpdateChannel.PERPETUAL_VL_2024,
        "description": "Project Professional 2024 (64-bit)",
    },
    "ltsc2024-full-x64": {
        "products": ["ProPlus2024Volume", "VisioPro2024Volume", "ProjectPro2024Volume"],
        "architecture": Architecture.X64,
        "channel": UpdateChannel.PERPETUAL_VL_2024,
        "description": "Office LTSC 2024 + Visio + Project (64-bit)",
    },
    "ltsc2024-full-x86": {
        "products": ["ProPlus2024Volume", "VisioPro2024Volume", "ProjectPro2024Volume"],
        "architecture": Architecture.X86,
        "channel": UpdateChannel.PERPETUAL_VL_2024,
        "description": "Office LTSC 2024 + Visio + Project (32-bit)",
    },
    "ltsc2021-full-x64": {
        "products": ["ProPlus2021Volume", "VisioPro2021Volume", "ProjectPro2021Volume"],
        "architecture": Architecture.X64,
        "channel": UpdateChannel.PERPETUAL_VL_2021,
        "description": "Office LTSC 2021 + Visio + Project (64-bit)",
    },
    "365-shared-computer": {
        "products": ["O365ProPlusRetail"],
        "architecture": Architecture.X64,
        "channel": UpdateChannel.CURRENT,
        "shared_computer": True,
        "description": "Microsoft 365 Apps with Shared Computer Licensing",
    },
    # Presets without OneDrive and Skype (Lync)
    "ltsc2024-full-x64-clean": {
        "products": ["ProPlus2024Volume", "VisioPro2024Volume", "ProjectPro2024Volume"],
        "architecture": Architecture.X64,
        "channel": UpdateChannel.PERPETUAL_VL_2024,
        "exclude_apps": ["OneDrive", "Lync"],
        "description": "Office LTSC 2024 + Visio + Project (64-bit) - No OneDrive/Skype",
    },
    "ltsc2024-full-x86-clean": {
        "products": ["ProPlus2024Volume", "VisioPro2024Volume", "ProjectPro2024Volume"],
        "architecture": Architecture.X86,
        "channel": UpdateChannel.PERPETUAL_VL_2024,
        "exclude_apps": ["OneDrive", "Lync"],
        "description": "Office LTSC 2024 + Visio + Project (32-bit) - No OneDrive/Skype",
    },
    "ltsc2024-x64-clean": {
        "products": ["ProPlus2024Volume"],
        "architecture": Architecture.X64,
        "channel": UpdateChannel.PERPETUAL_VL_2024,
        "exclude_apps": ["OneDrive", "Lync"],
        "description": "Office LTSC 2024 Professional Plus (64-bit) - No OneDrive/Skype",
    },
    "365-proplus-x64-clean": {
        "products": ["O365ProPlusRetail"],
        "architecture": Architecture.X64,
        "channel": UpdateChannel.CURRENT,
        "exclude_apps": ["OneDrive", "Lync"],
        "description": "Microsoft 365 Apps (64-bit) - No OneDrive/Skype",
    },
    "365-proplus-visio-project-clean": {
        "products": ["O365ProPlusRetail", "VisioProRetail", "ProjectProRetail"],
        "architecture": Architecture.X64,
        "channel": UpdateChannel.CURRENT,
        "exclude_apps": ["OneDrive", "Lync"],
        "description": "Microsoft 365 Apps + Visio + Project (64-bit) - No OneDrive/Skype",
    },
}
"""!
@brief Pre-built installation presets matching OEM config files.
"""


# ---------------------------------------------------------------------------
# Data Classes
# ---------------------------------------------------------------------------


@dataclass
class ProductConfig:
    """!
    @brief Configuration for a single Office product to install.
    """

    product_id: str
    languages: list[str] = field(default_factory=lambda: ["en-us"])
    exclude_apps: list[str] = field(default_factory=list)

    def validate(self) -> list[str]:
        """!
        @brief Validate the product configuration.
        @returns List of validation error messages (empty if valid).
        """
        errors: list[str] = []
        if self.product_id not in PRODUCT_IDS:
            errors.append(f"Unknown product ID: {self.product_id}")
        for lang in self.languages:
            if lang.lower() not in SUPPORTED_LANGUAGES:
                errors.append(f"Unsupported language: {lang}")
        return errors


@dataclass
class ODTConfig:
    """!
    @brief Complete ODT XML configuration for Office installation.
    """

    products: list[ProductConfig] = field(default_factory=list)
    architecture: Architecture = Architecture.X64
    channel: UpdateChannel = UpdateChannel.CURRENT
    accept_eula: bool = True
    display_level: DisplayLevel = DisplayLevel.NONE
    force_app_shutdown: bool = True
    pin_icons_to_taskbar: bool = False
    auto_activate: bool = True
    enable_updates: bool = True
    logging_path: str = "%temp%"
    logging_level: str = "Standard"
    source_path: str | None = None
    download_path: str | None = None
    shared_computer_licensing: bool = False
    device_based_licensing: bool = False
    scc_cache_override: str | None = None
    version: str | None = None
    allow_cdn_fallback: bool = True
    remove_msi: bool = False

    def validate(self) -> list[str]:
        """!
        @brief Validate the entire ODT configuration.
        @returns List of validation error messages (empty if valid).
        """
        errors: list[str] = []
        if not self.products:
            errors.append("At least one product must be specified")
        for product in self.products:
            errors.extend(product.validate())
        return errors

    @classmethod
    def from_preset(cls, preset_name: str, languages: list[str] | None = None) -> ODTConfig:
        """!
        @brief Create an ODT configuration from a preset name.
        @param preset_name Name of the preset from INSTALL_PRESETS.
        @param languages Optional list of languages (default: ["en-us"]).
        @returns Configured ODTConfig instance.
        @raises ValueError if preset is not found.
        """
        preset = INSTALL_PRESETS.get(preset_name)
        if not preset:
            available = ", ".join(sorted(INSTALL_PRESETS.keys()))
            raise ValueError(f"Unknown preset '{preset_name}'. Available: {available}")

        langs = languages or ["en-us"]
        exclude_apps = preset.get("exclude_apps", [])  # type: ignore[union-attr]
        products = [
            ProductConfig(product_id=pid, languages=langs, exclude_apps=list(exclude_apps))
            for pid in preset.get("products", [])  # type: ignore[union-attr]
        ]

        config = cls(
            products=products,
            architecture=preset.get("architecture", Architecture.X64),  # type: ignore[arg-type]
            channel=preset.get("channel", UpdateChannel.CURRENT),  # type: ignore[arg-type]
            shared_computer_licensing=preset.get("shared_computer", False),  # type: ignore[arg-type]
        )
        return config


# ---------------------------------------------------------------------------
# XML Builder
# ---------------------------------------------------------------------------


def _indent_xml(elem: ET.Element, level: int = 0) -> None:
    """!
    @brief Add indentation to XML elements for pretty printing.
    @param elem Root element to indent.
    @param level Current indentation level.
    """
    indent = "\n" + "  " * level
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = indent + "  "
        if not elem.tail or not elem.tail.strip():
            elem.tail = indent
        for child in elem:
            _indent_xml(child, level + 1)
        if not child.tail or not child.tail.strip():
            child.tail = indent
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = indent


def build_xml(config: ODTConfig) -> str:
    """!
    @brief Generate ODT XML configuration string from ODTConfig.
    @param config ODTConfig instance with installation parameters.
    @returns XML configuration as a string.
    @raises ValueError if configuration is invalid.
    """
    errors = config.validate()
    if errors:
        raise ValueError(f"Invalid configuration: {'; '.join(errors)}")

    # Create root Configuration element
    root = ET.Element("Configuration")

    # Add element
    add_elem = ET.SubElement(root, "Add")
    add_elem.set("OfficeClientEdition", config.architecture.value)
    add_elem.set("Channel", config.channel.value)

    if config.source_path:
        add_elem.set("SourcePath", config.source_path)
    if config.version:
        add_elem.set("Version", config.version)
    if config.allow_cdn_fallback:
        add_elem.set("AllowCdnFallback", "TRUE")

    # Add products
    for product in config.products:
        product_elem = ET.SubElement(add_elem, "Product")
        product_elem.set("ID", product.product_id)

        # Add languages
        for lang in product.languages:
            lang_elem = ET.SubElement(product_elem, "Language")
            lang_elem.set("ID", lang)

        # Add excluded apps
        for app in product.exclude_apps:
            exclude_elem = ET.SubElement(product_elem, "ExcludeApp")
            exclude_elem.set("ID", app)

    # Updates element
    updates_elem = ET.SubElement(root, "Updates")
    updates_elem.set("Enabled", "TRUE" if config.enable_updates else "FALSE")

    # Display element
    display_elem = ET.SubElement(root, "Display")
    display_elem.set("Level", config.display_level.value)
    display_elem.set("AcceptEULA", "TRUE" if config.accept_eula else "FALSE")

    # Logging element
    logging_elem = ET.SubElement(root, "Logging")
    logging_elem.set("Level", config.logging_level)
    logging_elem.set("Path", config.logging_path)

    # Property elements
    if config.force_app_shutdown:
        prop = ET.SubElement(root, "Property")
        prop.set("Name", "FORCEAPPSHUTDOWN")
        prop.set("Value", "TRUE")

    if not config.pin_icons_to_taskbar:
        prop = ET.SubElement(root, "Property")
        prop.set("Name", "PinIconsToTaskbar")
        prop.set("Value", "FALSE")

    if config.auto_activate:
        prop = ET.SubElement(root, "Property")
        prop.set("Name", "AUTOACTIVATE")
        prop.set("Value", "1")

    if config.shared_computer_licensing:
        prop = ET.SubElement(root, "Property")
        prop.set("Name", "SharedComputerLicensing")
        prop.set("Value", "1")

    if config.device_based_licensing:
        prop = ET.SubElement(root, "Property")
        prop.set("Name", "DeviceBasedLicensing")
        prop.set("Value", "1")

    if config.scc_cache_override:
        prop = ET.SubElement(root, "Property")
        prop.set("Name", "SCLCacheOverride")
        prop.set("Value", config.scc_cache_override)

    # RemoveMSI element
    if config.remove_msi:
        ET.SubElement(root, "RemoveMSI")

    # Pretty print
    _indent_xml(root)

    # Generate XML string with declaration
    xml_str = ET.tostring(root, encoding="unicode")
    return f'<?xml version="1.0" encoding="UTF-8"?>\n{xml_str}'


def build_removal_xml(
    *,
    remove_all: bool = True,
    product_ids: Sequence[str] | None = None,
    force_app_shutdown: bool = True,
    remove_msi: bool = True,
    display_level: DisplayLevel = DisplayLevel.NONE,
) -> str:
    """!
    @brief Generate ODT XML for Office removal.
    @param remove_all Remove all Office products (default: True).
    @param product_ids Specific product IDs to remove (if not remove_all).
    @param force_app_shutdown Force close running Office apps.
    @param remove_msi Also remove MSI-based Office installations.
    @param display_level UI visibility during removal.
    @returns XML configuration string for removal.
    """
    root = ET.Element("Configuration")

    # Remove element
    remove_elem = ET.SubElement(root, "Remove")
    if remove_all:
        remove_elem.set("All", "TRUE")
    elif product_ids:
        for pid in product_ids:
            product_elem = ET.SubElement(remove_elem, "Product")
            product_elem.set("ID", pid)

    # Property for force app shutdown
    if force_app_shutdown:
        prop = ET.SubElement(root, "Property")
        prop.set("Name", "FORCEAPPSHUTDOWN")
        prop.set("Value", "TRUE")

    # RemoveMSI element
    if remove_msi:
        ET.SubElement(root, "RemoveMSI")

    # Display element
    display_elem = ET.SubElement(root, "Display")
    display_elem.set("Level", display_level.value)
    display_elem.set("AcceptEULA", "TRUE")

    # Logging
    logging_elem = ET.SubElement(root, "Logging")
    logging_elem.set("Level", "Standard")
    logging_elem.set("Path", "%temp%")

    _indent_xml(root)
    xml_str = ET.tostring(root, encoding="unicode")
    return f'<?xml version="1.0" encoding="UTF-8"?>\n{xml_str}'


def build_download_xml(
    config: ODTConfig,
    download_path: str,
) -> str:
    """!
    @brief Generate ODT XML for downloading Office installation files.
    @param config ODTConfig with products and settings.
    @param download_path Local path to store downloaded files.
    @returns XML configuration string for download.
    """
    # Clone config and set download path
    download_config = ODTConfig(
        products=config.products,
        architecture=config.architecture,
        channel=config.channel,
        accept_eula=config.accept_eula,
        version=config.version,
    )

    errors = download_config.validate()
    if errors:
        raise ValueError(f"Invalid configuration: {'; '.join(errors)}")

    root = ET.Element("Configuration")

    # Add element with download path
    add_elem = ET.SubElement(root, "Add")
    add_elem.set("OfficeClientEdition", download_config.architecture.value)
    add_elem.set("Channel", download_config.channel.value)
    add_elem.set("SourcePath", download_path)

    if download_config.version:
        add_elem.set("Version", download_config.version)

    # Add products
    for product in download_config.products:
        product_elem = ET.SubElement(add_elem, "Product")
        product_elem.set("ID", product.product_id)
        for lang in product.languages:
            lang_elem = ET.SubElement(product_elem, "Language")
            lang_elem.set("ID", lang)

    _indent_xml(root)
    xml_str = ET.tostring(root, encoding="unicode")
    return f'<?xml version="1.0" encoding="UTF-8"?>\n{xml_str}'


# ---------------------------------------------------------------------------
# File Operations
# ---------------------------------------------------------------------------


def write_xml_config(config: ODTConfig, output_path: str | Path) -> Path:
    """!
    @brief Write ODT XML configuration to a file.
    @param config ODTConfig instance.
    @param output_path Destination file path.
    @returns Path to the written file.
    """
    log = logging_ext.get_human_logger()
    output_path = Path(output_path)

    xml_content = build_xml(config)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(xml_content, encoding="utf-8")

    log.info(f"ODT configuration written to: {output_path}")
    return output_path


def write_temp_config(config: ODTConfig, prefix: str = "odt_install_") -> Path:
    """!
    @brief Write ODT XML to a temporary file.
    @param config ODTConfig instance.
    @param prefix Filename prefix for the temp file.
    @returns Path to the temporary file.
    """
    xml_content = build_xml(config)

    fd, temp_path = tempfile.mkstemp(suffix=".xml", prefix=prefix)
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(xml_content)
    except Exception:
        os.close(fd)
        raise

    return Path(temp_path)


# ---------------------------------------------------------------------------
# ODT Execution
# ---------------------------------------------------------------------------


def get_odt_setup_path() -> Path:
    """!
    @brief Get path to the embedded ODT setup.exe.
    @details Looks for setup.exe in the oem/ folder relative to this module,
    or in the PyInstaller _MEIPASS temp directory when running as frozen exe.
    @returns Path to setup.exe.
    @raises FileNotFoundError if setup.exe cannot be found.
    """
    # When running as PyInstaller bundle
    if getattr(sys, "frozen", False):
        base_path = Path(sys._MEIPASS)  # type: ignore[attr-defined]
    else:
        # Development: look relative to this module
        base_path = Path(__file__).parent.parent.parent

    oem_path = base_path / "oem" / "setup.exe"
    if oem_path.exists():
        return oem_path

    # Also check directly in base (for alternative packaging)
    alt_path = base_path / "setup.exe"
    if alt_path.exists():
        return alt_path

    raise FileNotFoundError(f"ODT setup.exe not found. Checked:\n  {oem_path}\n  {alt_path}")


@dataclass
class ODTResult:
    """!
    @brief Result from an ODT operation.
    """

    success: bool
    return_code: int
    command: list[str]
    config_path: Path | None
    stdout: str
    stderr: str
    duration: float
    error: str | None = None


# ---------------------------------------------------------------------------
# ODT Progress Monitoring
# ---------------------------------------------------------------------------


# Common Office installation paths to monitor
_OFFICE_INSTALL_PATHS = [
    Path(os.environ.get("ProgramFiles", "C:\\Program Files")) / "Microsoft Office",
    Path(os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)")) / "Microsoft Office",
    Path(os.environ.get("ProgramFiles", "C:\\Program Files"))
    / "Common Files"
    / "microsoft shared"
    / "ClickToRun",
    Path(os.environ.get("ProgramData", "C:\\ProgramData")) / "Microsoft" / "ClickToRun",
]

# Registry keys to monitor for Office installation
_OFFICE_REGISTRY_KEYS = [
    r"SOFTWARE\Microsoft\Office\ClickToRun",
    r"SOFTWARE\Microsoft\Office\16.0",
    r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
]


def _get_odt_log_path() -> Path:
    """!
    @brief Get the path where ODT writes installation logs.
    @returns Path to the ODT log directory (typically %TEMP%).
    """
    return Path(os.environ.get("TEMP", tempfile.gettempdir()))


def _find_latest_odt_log() -> Path | None:
    """!
    @brief Find the most recent ODT Click-to-Run log file.
    @returns Path to the latest log file, or None if not found.
    """
    log_dir = _get_odt_log_path()
    # ODT creates logs like "Microsoft Office Click-to-Run*.log"
    pattern = str(log_dir / "Microsoft Office Click-to-Run*.log")
    logs = glob.glob(pattern)
    if not logs:
        return None
    # Return the most recently modified
    return Path(max(logs, key=os.path.getmtime))


def _get_folder_size(path: Path) -> int:
    """!
    @brief Get total size of a folder in bytes.
    @details Yields GIL every 100 files to keep spinner responsive.
    @param path Path to the folder.
    @returns Total size in bytes, or 0 if folder doesn't exist.
    """
    if not path.exists():
        return 0
    try:
        total = 0
        count = 0
        for entry in path.rglob("*"):
            if entry.is_file():
                try:
                    total += entry.stat().st_size
                except (OSError, PermissionError):
                    pass
            # Yield GIL every 100 files to keep spinner responsive
            count += 1
            if count % 100 == 0:
                time.sleep(0)
        return total
    except (OSError, PermissionError):
        return 0


def _get_office_install_size() -> int:
    """!
    @brief Get combined size of all Office installation folders.
    @returns Total size in bytes.
    """
    total = 0
    for path in _OFFICE_INSTALL_PATHS:
        total += _get_folder_size(path)
    return total


def _count_office_files() -> int:
    """!
    @brief Count files in Office installation folders.
    @details Yields GIL every 100 files to keep spinner responsive.
    @returns Total file count.
    """
    total = 0
    yield_counter = 0
    for path in _OFFICE_INSTALL_PATHS:
        if path.exists():
            try:
                for entry in path.rglob("*"):
                    if entry.is_file():
                        total += 1
                    yield_counter += 1
                    if yield_counter % 100 == 0:
                        time.sleep(0)  # Yield GIL
            except (OSError, PermissionError):
                pass
    return total


def _check_registry_key_exists(key_path: str) -> bool:
    """!
    @brief Check if a registry key exists under HKLM.
    @param key_path Registry key path (without HKLM prefix).
    @returns True if key exists.
    """
    if sys.platform != "win32":
        return False
    try:
        import winreg

        winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path)
        return True
    except (FileNotFoundError, OSError):
        return False


def _count_registry_subkeys(key_path: str) -> int:
    """!
    @brief Count subkeys under a registry key.
    @param key_path Registry key path (without HKLM prefix).
    @returns Number of subkeys, or 0 if key doesn't exist.
    """
    if sys.platform != "win32":
        return 0
    try:
        import winreg

        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path)
        info = winreg.QueryInfoKey(key)
        winreg.CloseKey(key)
        return info[0]  # Number of subkeys
    except (FileNotFoundError, OSError):
        return 0


def _get_c2r_version() -> str | None:
    """!
    @brief Get the installed Click-to-Run version from registry.
    @returns Version string or None if not found.
    """
    if sys.platform != "win32":
        return None
    try:
        import winreg

        key = winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
        )
        version, _ = winreg.QueryValueEx(key, "VersionToReport")
        winreg.CloseKey(key)
        return str(version)
    except (FileNotFoundError, OSError, ValueError):
        return None


def _format_size(size_bytes: int) -> str:
    """!
    @brief Format byte size as human-readable string.
    @param size_bytes Size in bytes.
    @returns Formatted string like "1.5 GB" or "256 MB".
    """
    if size_bytes >= 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024 * 1024):.1f} GB"
    elif size_bytes >= 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.0f} MB"
    elif size_bytes >= 1024:
        return f"{size_bytes / 1024:.0f} KB"
    return f"{size_bytes} B"


# ClickToRun process names to monitor after setup.exe exits
_CLICKTORUN_PROCESS_NAMES = {
    "officeclicktorun.exe",
    "officec2rclient.exe",
    "officeservicemanager.exe",
    "integratedoffice.exe",
    "appvshnotify.exe",
}


def _find_running_clicktorun_processes() -> list[tuple[int, str]]:
    """!
    @brief Find running ClickToRun-related processes.
    @returns List of (pid, process_name) tuples for running C2R processes.
    """
    if sys.platform != "win32":
        return []

    result = []
    try:
        # Use tasklist to get running processes
        proc = subprocess.run(
            ["tasklist", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            timeout=10,
        )
        if proc.returncode == 0:
            for line in proc.stdout.strip().split("\n"):
                if not line:
                    continue
                # Parse CSV: "process.exe","PID","Session","Session#","Mem Usage"
                parts = line.split('","')
                if len(parts) >= 2:
                    proc_name = parts[0].strip('"').lower()
                    if proc_name in _CLICKTORUN_PROCESS_NAMES:
                        try:
                            pid = int(parts[1].strip('"'))
                            result.append((pid, proc_name))
                        except ValueError:
                            pass
    except Exception:
        pass

    return result


def _wait_for_clicktorun_processes(
    stop_event: threading.Event,
    timeout: float | None = None,
) -> bool:
    """!
    @brief Wait for all ClickToRun processes to finish.
    @param stop_event Event to signal early termination.
    @param timeout Maximum seconds to wait (None = no limit).
    @returns True if all processes finished, False if timed out or interrupted.
    """
    start = time.monotonic()
    while not stop_event.is_set():
        running = _find_running_clicktorun_processes()
        if not running:
            return True

        if timeout and (time.monotonic() - start) > timeout:
            return False

        time.sleep(1.0)  # Check every second

    return False


def _get_process_cpu_percent(pid: int, interval: float = 0.1) -> float:
    """!
    @brief Get CPU usage percentage for a process using WMI.
    @param pid Process ID to monitor.
    @param interval Sampling interval in seconds.
    @returns CPU percentage (0-100 per core), or 0.0 if unavailable.
    """
    if sys.platform != "win32":
        return 0.0
    try:
        import ctypes
        from ctypes import wintypes

        # Get process handle
        PROCESS_QUERY_INFORMATION = 0x0400
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.OpenProcess(PROCESS_QUERY_INFORMATION, False, pid)
        if not handle:
            return 0.0

        try:
            # FILETIME structure
            class FILETIME(ctypes.Structure):
                _fields_ = [("dwLowDateTime", wintypes.DWORD), ("dwHighDateTime", wintypes.DWORD)]

            creation_time = FILETIME()
            exit_time = FILETIME()
            kernel_time1 = FILETIME()
            user_time1 = FILETIME()
            kernel_time2 = FILETIME()
            user_time2 = FILETIME()

            # First measurement
            if not kernel32.GetProcessTimes(
                handle,
                ctypes.byref(creation_time),
                ctypes.byref(exit_time),
                ctypes.byref(kernel_time1),
                ctypes.byref(user_time1),
            ):
                return 0.0

            time.sleep(interval)

            # Second measurement
            if not kernel32.GetProcessTimes(
                handle,
                ctypes.byref(creation_time),
                ctypes.byref(exit_time),
                ctypes.byref(kernel_time2),
                ctypes.byref(user_time2),
            ):
                return 0.0

            # Calculate CPU time delta (100ns units)
            def filetime_to_int(ft: FILETIME) -> int:
                return (ft.dwHighDateTime << 32) | ft.dwLowDateTime

            kernel_delta = filetime_to_int(kernel_time2) - filetime_to_int(kernel_time1)
            user_delta = filetime_to_int(user_time2) - filetime_to_int(user_time1)
            total_delta = kernel_delta + user_delta

            # Convert to percentage (100ns to seconds, then to percent)
            # interval is in seconds, FILETIME is 100ns units
            cpu_percent = (total_delta / (interval * 10_000_000)) * 100
            return min(cpu_percent, 100.0 * os.cpu_count())  # Cap at max CPU

        finally:
            kernel32.CloseHandle(handle)

    except Exception:
        return 0.0


def _get_process_memory_mb(pid: int) -> float:
    """!
    @brief Get memory usage in MB for a process.
    @param pid Process ID to monitor.
    @returns Memory usage in MB, or 0.0 if unavailable.
    """
    if sys.platform != "win32":
        return 0.0
    try:
        import ctypes
        from ctypes import wintypes

        # Get process handle
        PROCESS_QUERY_INFORMATION = 0x0400
        PROCESS_VM_READ = 0x0010
        kernel32 = ctypes.windll.kernel32
        handle = kernel32.OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, False, pid)
        if not handle:
            return 0.0

        try:
            # PROCESS_MEMORY_COUNTERS structure
            class PROCESS_MEMORY_COUNTERS(ctypes.Structure):
                _fields_ = [
                    ("cb", wintypes.DWORD),
                    ("PageFaultCount", wintypes.DWORD),
                    ("PeakWorkingSetSize", ctypes.c_size_t),
                    ("WorkingSetSize", ctypes.c_size_t),
                    ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
                    ("QuotaPagedPoolUsage", ctypes.c_size_t),
                    ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
                    ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
                    ("PagefileUsage", ctypes.c_size_t),
                    ("PeakPagefileUsage", ctypes.c_size_t),
                ]

            psapi = ctypes.windll.psapi
            counters = PROCESS_MEMORY_COUNTERS()
            counters.cb = ctypes.sizeof(PROCESS_MEMORY_COUNTERS)

            if psapi.GetProcessMemoryInfo(handle, ctypes.byref(counters), counters.cb):
                return counters.WorkingSetSize / (1024 * 1024)
            return 0.0

        finally:
            kernel32.CloseHandle(handle)

    except Exception:
        return 0.0


def _get_process_tree_stats(pid: int) -> tuple[float, float]:
    """!
    @brief Get combined CPU and memory for a process and its children.
    @param pid Root process ID.
    @returns Tuple of (cpu_percent, memory_mb) for the process tree.
    """
    if sys.platform != "win32":
        return 0.0, 0.0

    try:
        import ctypes

        # Get all child PIDs using CreateToolhelp32Snapshot
        TH32CS_SNAPPROCESS = 0x00000002
        kernel32 = ctypes.windll.kernel32

        class PROCESSENTRY32(ctypes.Structure):
            _fields_ = [
                ("dwSize", ctypes.c_ulong),
                ("cntUsage", ctypes.c_ulong),
                ("th32ProcessID", ctypes.c_ulong),
                ("th32DefaultHeapID", ctypes.c_void_p),
                ("th32ModuleID", ctypes.c_ulong),
                ("cntThreads", ctypes.c_ulong),
                ("th32ParentProcessID", ctypes.c_ulong),
                ("pcPriClassBase", ctypes.c_long),
                ("dwFlags", ctypes.c_ulong),
                ("szExeFile", ctypes.c_char * 260),
            ]

        snapshot = kernel32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
        if snapshot == -1:
            return _get_process_cpu_percent(pid, 0.1), _get_process_memory_mb(pid)

        try:
            # Find all child processes
            pids_to_check = {pid}
            all_pids = set()

            pe32 = PROCESSENTRY32()
            pe32.dwSize = ctypes.sizeof(PROCESSENTRY32)

            if kernel32.Process32First(snapshot, ctypes.byref(pe32)):
                while True:
                    all_pids.add((pe32.th32ProcessID, pe32.th32ParentProcessID))
                    if not kernel32.Process32Next(snapshot, ctypes.byref(pe32)):
                        break

            # Build tree of child PIDs
            found_new = True
            while found_new:
                found_new = False
                for child_pid, parent_pid in all_pids:
                    if parent_pid in pids_to_check and child_pid not in pids_to_check:
                        pids_to_check.add(child_pid)
                        found_new = True

        finally:
            kernel32.CloseHandle(snapshot)

        # Sum up stats for all processes in tree
        total_memory = 0.0
        for p in pids_to_check:
            total_memory += _get_process_memory_mb(p)

        # CPU sampling (just for root process to avoid long delay)
        cpu = _get_process_cpu_percent(pid, 0.05)

        return cpu, total_memory

    except Exception:
        return 0.0, 0.0


def _parse_odt_progress(log_path: Path) -> tuple[str, int | None]:
    """!
    @brief Parse ODT log file for installation progress.
    @param log_path Path to the ODT log file.
    @returns Tuple of (status_message, percentage or None).
    """
    if not log_path.exists():
        return "Starting installation...", None

    try:
        # Read last portion of log (it can be large)
        with open(log_path, encoding="utf-8", errors="ignore") as f:
            # Seek to end and read last 8KB
            f.seek(0, 2)  # End of file
            size = f.tell()
            f.seek(max(0, size - 8192))
            content = f.read()

        lines = content.strip().split("\n")

        # Look for progress indicators from bottom up
        for line in reversed(lines[-50:]):  # Check last 50 lines
            line_lower = line.lower()

            # Check for percentage patterns
            pct_match = re.search(r"(\d{1,3})%", line)

            # Check for specific status messages
            if "downloading" in line_lower:
                pct = int(pct_match.group(1)) if pct_match else None
                return "Downloading Office files", pct
            elif "installing" in line_lower:
                pct = int(pct_match.group(1)) if pct_match else None
                return "Installing Office", pct
            elif "configuring" in line_lower:
                return "Configuring Office", None
            elif "applying" in line_lower:
                return "Applying settings", None
            elif "finalizing" in line_lower or "completing" in line_lower:
                return "Finalizing installation", None
            elif "registering" in line_lower:
                return "Registering components", None
            elif "updating" in line_lower:
                return "Updating Office", None
            elif "removing" in line_lower:
                return "Removing old version", None
            elif "verifying" in line_lower:
                return "Verifying installation", None
            elif pct_match:
                # Found a percentage without specific context
                return "Installing Office", int(pct_match.group(1))

        return "Installing Office...", None

    except Exception:
        return "Installing Office...", None


@dataclass
class InstallMetrics:
    """!
    @brief Metrics captured during Office installation.
    """

    install_size: int = 0
    file_count: int = 0
    registry_keys: int = 0
    c2r_version: str | None = None
    log_status: str = ""
    log_percent: int | None = None
    cpu_percent: float = 0.0
    memory_mb: float = 0.0


@dataclass
class _MonitorStats:
    """!
    @brief Shared container for all monitor stats updated by polling thread.
    @details Caches all expensive metrics (disk, files, registry, CPU, RAM)
    to reduce resource usage - polling happens every 900ms instead of real-time.
    """

    # Process stats
    cpu_percent: float = 0.0
    memory_mb: float = 0.0

    # Disk/file stats
    install_size: int = 0
    file_count: int = 0
    registry_keys: int = 0

    # Log parsing stats
    log_status: str = ""
    log_percent: int | None = None

    # Download-specific stats
    download_size: int = 0
    download_files: int = 0

    # Thread safety
    _lock: threading.Lock = field(default_factory=threading.Lock)

    def update_install(
        self,
        cpu: float,
        mem: float,
        size: int,
        files: int,
        reg_keys: int,
        log_status: str,
        log_pct: int | None,
    ) -> None:
        """Thread-safe update of install stats."""
        with self._lock:
            self.cpu_percent = cpu
            self.memory_mb = mem
            self.install_size = size
            self.file_count = files
            self.registry_keys = reg_keys
            self.log_status = log_status
            self.log_percent = log_pct

    def update_download(
        self,
        cpu: float,
        mem: float,
        dl_size: int,
        dl_files: int,
        log_status: str,
        log_pct: int | None,
    ) -> None:
        """Thread-safe update of download stats."""
        with self._lock:
            self.cpu_percent = cpu
            self.memory_mb = mem
            self.download_size = dl_size
            self.download_files = dl_files
            self.log_status = log_status
            self.log_percent = log_pct

    def get_install(self) -> tuple[float, float, int, int, int, str, int | None]:
        """Thread-safe read of install stats."""
        with self._lock:
            return (
                self.cpu_percent,
                self.memory_mb,
                self.install_size,
                self.file_count,
                self.registry_keys,
                self.log_status,
                self.log_percent,
            )

    def get_download(self) -> tuple[float, float, int, int, str, int | None]:
        """Thread-safe read of download stats."""
        with self._lock:
            return (
                self.cpu_percent,
                self.memory_mb,
                self.download_size,
                self.download_files,
                self.log_status,
                self.log_percent,
            )


# Legacy alias for backward compatibility
_ProcessStats = _MonitorStats


def _install_poller_thread(
    pid: int,
    stats: _MonitorStats,
    stop_event: threading.Event,
    interval: float = 0.5,
) -> None:
    """!
    @brief Dedicated thread to poll install metrics without blocking spinner.
    @details Polls lightweight stats frequently, heavy disk/file ops less often.
    @param pid Process ID to monitor.
    @param stats Shared _MonitorStats object to update.
    @param stop_event Event to signal when polling should stop.
    @param interval Polling interval in seconds (default 500ms).
    """
    heavy_poll_counter = 0
    HEAVY_POLL_INTERVAL = 5  # Only poll disk/files every 5 intervals (~2.5s)

    # Cache for heavy operations
    cached_size = 0
    cached_files = 0
    cached_reg_keys = 0
    cached_cpu = 0.0
    cached_mem = 0.0

    while not stop_event.is_set():
        try:
            # Yield GIL at start of loop
            time.sleep(0)

            # Poll CPU/RAM - these can block briefly, so do them less often
            heavy_poll_counter += 1
            if heavy_poll_counter >= HEAVY_POLL_INTERVAL:
                heavy_poll_counter = 0

                # CPU measurement with short interval (blocks for 50ms)
                try:
                    cached_cpu = _get_process_cpu_percent(pid, interval=0.05)
                except Exception:
                    cached_cpu = 0.0

                time.sleep(0)  # Yield GIL

                try:
                    cached_mem = _get_process_memory_mb(pid)
                except Exception:
                    cached_mem = 0.0

                time.sleep(0)  # Yield GIL

                # Heavy disk/file operations
                try:
                    cached_size = _get_office_install_size()
                except Exception:
                    pass

                time.sleep(0)  # Yield GIL

                try:
                    cached_files = _count_office_files()
                except Exception:
                    pass

                time.sleep(0)  # Yield GIL

                try:
                    cached_reg_keys = sum(
                        _count_registry_subkeys(key) for key in _OFFICE_REGISTRY_KEYS
                    )
                except Exception:
                    pass

            # Log status is lightweight - check every poll
            log_status, log_pct = "Starting...", None
            try:
                log_path = _find_latest_odt_log()
                if log_path:
                    log_status, log_pct = _parse_odt_progress(log_path)
            except Exception:
                pass

            # Update shared stats
            stats.update_install(
                cached_cpu,
                cached_mem,
                cached_size,
                cached_files,
                cached_reg_keys,
                log_status,
                log_pct,
            )

        except Exception:
            pass

        # Wait for next poll - short interval for responsiveness
        stop_event.wait(interval)


def _download_poller_thread(
    pid: int,
    download_path: Path,
    stats: _MonitorStats,
    stop_event: threading.Event,
    interval: float = 0.5,
) -> None:
    """!
    @brief Dedicated thread to poll download metrics without blocking spinner.
    @details Polls lightweight stats frequently, heavy disk ops less often.
    @param pid Process ID to monitor.
    @param download_path Path where files are being downloaded.
    @param stats Shared _MonitorStats object to update.
    @param stop_event Event to signal when polling should stop.
    @param interval Polling interval in seconds (default 500ms).
    """
    heavy_poll_counter = 0
    HEAVY_POLL_INTERVAL = 5  # Only poll disk every 5 intervals (~2.5s)

    # Cache for heavy operations
    cached_dl_size = 0
    cached_dl_files = 0
    cached_cpu = 0.0
    cached_mem = 0.0

    while not stop_event.is_set():
        try:
            # Yield GIL at start of loop
            time.sleep(0)

            # Poll CPU/RAM and disk - do them together less often
            heavy_poll_counter += 1
            if heavy_poll_counter >= HEAVY_POLL_INTERVAL:
                heavy_poll_counter = 0

                try:
                    cached_cpu = _get_process_cpu_percent(pid, interval=0.05)
                except Exception:
                    cached_cpu = 0.0

                time.sleep(0)  # Yield GIL

                try:
                    cached_mem = _get_process_memory_mb(pid)
                except Exception:
                    cached_mem = 0.0

                time.sleep(0)  # Yield GIL

                try:
                    cached_dl_size = _get_folder_size(download_path)
                except Exception:
                    pass

                time.sleep(0)  # Yield GIL

                try:
                    cached_dl_files = sum(1 for _ in download_path.rglob("*") if _.is_file())
                except Exception:
                    pass

            # Log status is lightweight - check every poll
            log_status, log_pct = "Starting...", None
            try:
                log_path = _find_latest_odt_log()
                if log_path:
                    log_status, log_pct = _parse_odt_progress(log_path)
            except Exception:
                pass

            # Update shared stats
            stats.update_download(
                cached_cpu, cached_mem, cached_dl_size, cached_dl_files, log_status, log_pct
            )

        except Exception:
            pass

        # Wait for next poll - short interval for responsiveness
        stop_event.wait(interval)


def _capture_install_metrics(
    pid: int | None = None,
    cached_stats: _MonitorStats | None = None,
) -> InstallMetrics:
    """!
    @brief Capture current installation metrics.
    @param pid Process ID to monitor for CPU/RAM (optional).
    @param cached_stats Pre-polled _MonitorStats object (preferred for CPU/RAM).
    @returns InstallMetrics with current state.
    """
    metrics = InstallMetrics()
    metrics.install_size = _get_office_install_size()
    metrics.file_count = _count_office_files()
    metrics.registry_keys = sum(_count_registry_subkeys(key) for key in _OFFICE_REGISTRY_KEYS)
    metrics.c2r_version = _get_c2r_version()

    log_path = _find_latest_odt_log()
    if log_path:
        metrics.log_status, metrics.log_percent = _parse_odt_progress(log_path)
    else:
        metrics.log_status = "Starting..."

    # Get process stats from cached poller (non-blocking) or direct call
    if cached_stats is not None:
        # Use cached CPU/RAM from MonitorStats
        cpu, mem, _, _, _, _, _ = cached_stats.get_install()
        metrics.cpu_percent = cpu
        metrics.memory_mb = mem
    elif pid is not None:
        # Fallback: direct call (may block briefly)
        metrics.cpu_percent, metrics.memory_mb = _get_process_tree_stats(pid)

    return metrics


def _monitor_odt_progress(
    proc: subprocess.Popen[str],
    stop_event: threading.Event,
    products: list[str],
) -> None:
    """!
    @brief Background thread to monitor ODT installation progress.
    @details Uses a dedicated poller thread for I/O, this thread only updates spinner.
    @param proc The ODT subprocess being monitored.
    @param stop_event Event to signal when monitoring should stop.
    @param products List of product names being installed.
    """
    if _spinner is None:
        return

    product_str = ", ".join(products[:2])
    if len(products) > 2:
        product_str += f" +{len(products) - 2}"

    # Shared stats object for non-blocking reads
    stats = _MonitorStats()
    poller_stop = threading.Event()

    # Start dedicated poller thread for all I/O operations
    pid = proc.pid
    poller = threading.Thread(
        target=_install_poller_thread,
        args=(pid, stats, poller_stop),
        daemon=True,
        name="odt-install-poller",
    )
    poller.start()

    try:
        # Update spinner with stats every 200ms (non-blocking reads only)
        while not stop_event.is_set() and proc.poll() is None:
            # Read cached stats (non-blocking, thread-safe)
            cpu, mem, size_bytes, files, reg_keys, log_status, log_pct = stats.get_install()

            # Format size
            if size_bytes >= 1_000_000_000:
                size_str = f"{size_bytes / 1_000_000_000:.1f}GB"
            elif size_bytes >= 1_000_000:
                size_str = f"{size_bytes / 1_000_000:.0f}MB"
            else:
                size_str = f"{size_bytes / 1_000:.0f}KB"

            # Build status line
            parts = [f"ODT: {product_str}"]
            if log_pct is not None:
                parts.append(f"{log_pct}%")
            if log_status and log_status != "Starting...":
                # Truncate long status
                status_short = log_status[:30] + "..." if len(log_status) > 33 else log_status
                parts.append(status_short)

            # Add metrics
            metrics = []
            if size_bytes > 0:
                metrics.append(size_str)
            if files > 0:
                metrics.append(f"{files} files")
            if reg_keys > 0:
                metrics.append(f"{reg_keys} keys")
            if cpu > 0:
                metrics.append(f"CPU {cpu:.0f}%")
            if mem > 0:
                metrics.append(f"RAM {mem:.0f}MB")

            if metrics:
                parts.append(f"[{', '.join(metrics)}]")

            # Update spinner text (non-blocking)
            _spinner.update_task(" ".join(parts))

            # Short sleep to keep responsive
            stop_event.wait(0.2)
    finally:
        # Stop the poller thread
        poller_stop.set()
        poller.join(timeout=1.0)


def _monitor_clicktorun_progress(
    stop_event: threading.Event,
    products: list[str],
) -> None:
    """!
    @brief Background thread to monitor ClickToRun process progress.
    @details Used when setup.exe exits but C2R processes are still running.
    Monitors disk usage, file count, registry keys, and C2R process stats.
    @param stop_event Event to signal when monitoring should stop.
    @param products List of product names being installed.
    """
    if _spinner is None:
        return

    product_str = ", ".join(products[:2])
    if len(products) > 2:
        product_str += f" +{len(products) - 2}"

    # Shared stats object for non-blocking reads
    stats = _MonitorStats()
    poller_stop = threading.Event()

    # Start dedicated poller thread - use PID 0 to signal "find C2R processes"
    poller = threading.Thread(
        target=_clicktorun_poller_thread,
        args=(stats, poller_stop),
        daemon=True,
        name="c2r-poller",
    )
    poller.start()

    try:
        # Update spinner with stats every 200ms (non-blocking reads only)
        while not stop_event.is_set():
            # Check if any C2R processes are still running
            c2r_procs = _find_running_clicktorun_processes()
            if not c2r_procs:
                break

            # Read cached stats (non-blocking, thread-safe)
            cpu, mem, size_bytes, files, reg_keys, log_status, log_pct = stats.get_install()

            # Format size
            if size_bytes >= 1_000_000_000:
                size_str = f"{size_bytes / 1_000_000_000:.1f}GB"
            elif size_bytes >= 1_000_000:
                size_str = f"{size_bytes / 1_000_000:.0f}MB"
            else:
                size_str = f"{size_bytes / 1_000:.0f}KB"

            # Build status line showing C2R monitoring
            c2r_names = ", ".join(set(p[1].replace(".exe", "") for p in c2r_procs[:2]))
            parts = [f"C2R: {c2r_names}"]
            if log_pct is not None:
                parts.append(f"{log_pct}%")
            if log_status and log_status != "Starting...":
                status_short = log_status[:25] + "..." if len(log_status) > 28 else log_status
                parts.append(status_short)

            # Add metrics
            metrics = []
            if size_bytes > 0:
                metrics.append(size_str)
            if files > 0:
                metrics.append(f"{files} files")
            if reg_keys > 0:
                metrics.append(f"{reg_keys} keys")
            if cpu > 0:
                metrics.append(f"CPU {cpu:.0f}%")
            if mem > 0:
                metrics.append(f"RAM {mem:.0f}MB")

            if metrics:
                parts.append(f"[{', '.join(metrics)}]")

            # Update spinner text (non-blocking)
            _spinner.update_task(" ".join(parts))

            # Short sleep to keep responsive
            stop_event.wait(0.2)
    finally:
        # Stop the poller thread
        poller_stop.set()
        poller.join(timeout=1.0)


def _clicktorun_poller_thread(
    stats: _MonitorStats,
    stop_event: threading.Event,
    interval: float = 0.5,
) -> None:
    """!
    @brief Dedicated thread to poll ClickToRun process metrics.
    @details Monitors all running C2R processes for CPU/RAM, plus disk and registry.
    @param stats Shared _MonitorStats object to update.
    @param stop_event Event to signal when polling should stop.
    @param interval Polling interval in seconds.
    """
    heavy_poll_counter = 0
    HEAVY_POLL_INTERVAL = 5

    # Cache for heavy operations
    cached_size = 0
    cached_files = 0
    cached_reg_keys = 0
    cached_cpu = 0.0
    cached_mem = 0.0

    while not stop_event.is_set():
        try:
            time.sleep(0)  # Yield GIL

            heavy_poll_counter += 1
            if heavy_poll_counter >= HEAVY_POLL_INTERVAL:
                heavy_poll_counter = 0

                # Get stats from all running C2R processes
                c2r_procs = _find_running_clicktorun_processes()
                total_cpu = 0.0
                total_mem = 0.0
                for pid, _ in c2r_procs:
                    try:
                        total_cpu += _get_process_cpu_percent(pid, interval=0.05)
                    except Exception:
                        pass
                    time.sleep(0)
                    try:
                        total_mem += _get_process_memory_mb(pid)
                    except Exception:
                        pass
                cached_cpu = total_cpu
                cached_mem = total_mem

                time.sleep(0)

                try:
                    cached_size = _get_office_install_size()
                except Exception:
                    pass

                time.sleep(0)

                try:
                    cached_files = _count_office_files()
                except Exception:
                    pass

                time.sleep(0)

                try:
                    cached_reg_keys = sum(
                        _count_registry_subkeys(key) for key in _OFFICE_REGISTRY_KEYS
                    )
                except Exception:
                    pass

            # Log status - check every poll
            log_status, log_pct = "Installing...", None
            try:
                log_path = _find_latest_odt_log()
                if log_path:
                    log_status, log_pct = _parse_odt_progress(log_path)
            except Exception:
                pass

            stats.update_install(
                cached_cpu,
                cached_mem,
                cached_size,
                cached_files,
                cached_reg_keys,
                log_status,
                log_pct,
            )

        except Exception:
            pass

        stop_event.wait(interval)


def run_odt_install(
    config: ODTConfig,
    *,
    dry_run: bool = False,
    timeout: float | None = None,
    show_progress: bool = True,
) -> ODTResult:
    """!
    @brief Run ODT setup.exe to install Office with the given configuration.
    @param config ODTConfig with products and settings.
    @param dry_run If True, only generate the XML and print the command.
    @param timeout Optional timeout in seconds.
    @param show_progress If True, display spinner with progress updates.
    @returns ODTResult with execution details.
    """
    log = logging_ext.get_human_logger()

    try:
        setup_path = get_odt_setup_path()
    except FileNotFoundError as e:
        return ODTResult(
            success=False,
            return_code=-1,
            command=[],
            config_path=None,
            stdout="",
            stderr=str(e),
            duration=0.0,
            error=str(e),
        )

    # Write config to temp file
    config_path = write_temp_config(config, prefix="odt_install_")

    command = [str(setup_path), "/configure", str(config_path)]

    if dry_run:
        log.info(f"Installing Office with ODT: {' '.join(command)}")
        log.info(f"Config file: {config_path}")
        log.info("[DRY-RUN] Would execute: %s", " ".join(command))
        xml_content = config_path.read_text(encoding="utf-8")
        log.info(f"Configuration XML:\n{xml_content}")
        return ODTResult(
            success=True,
            return_code=0,
            command=command,
            config_path=config_path,
            stdout="",
            stderr="",
            duration=0.0,
        )

    # Get product names for progress display
    product_names = [p.product_id for p in config.products]

    # Log BEFORE starting spinner so messages appear correctly
    log.info(f"Installing Office with ODT: {' '.join(command)}")
    log.info(f"Config file: {config_path}")
    log.warning(
        "Office installation may take 10-30 minutes depending on products selected "
        "and network speed. Please be patient."
    )

    # Start spinner AFTER logging
    spinner_started = False
    if show_progress and (_spinner is not None):
        _spinner.start_spinner_thread()
        _spinner.set_task(f"ODT: Starting installation - {', '.join(product_names[:2])}")
        spinner_started = True

    start_time = time.monotonic()
    stop_event = threading.Event()
    monitor_thread: threading.Thread | None = None
    c2r_monitor_thread: threading.Thread | None = None
    proc: subprocess.Popen[str] | None = None

    try:
        # Start the ODT process
        # Don't capture stdout/stderr - ODT writes to its own logs
        # Using DEVNULL prevents pipe buffer blocking issues on Windows
        proc = subprocess.Popen(
            command,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            text=True,
        )

        # Register for cleanup on Ctrl+C
        if _spinner is not None:
            _spinner.register_process(proc)

        # Start progress monitoring thread
        if show_progress and (_spinner is not None):
            monitor_thread = threading.Thread(
                target=_monitor_odt_progress,
                args=(proc, stop_event, product_names),
                daemon=True,
                name="odt-progress",
            )
            monitor_thread.start()

        # Wait for process with polling to keep threads responsive
        return_code = None
        deadline = time.monotonic() + timeout if timeout else None
        while return_code is None:
            return_code = proc.poll()
            if return_code is None:
                if deadline and time.monotonic() > deadline:
                    raise subprocess.TimeoutExpired(command, timeout)
                time.sleep(0.1)  # Poll every 100ms

        stdout, stderr = "", ""
        error = None

        # Check if setup.exe exited but ClickToRun processes are still running
        # This happens when setup.exe exits early (error or handoff) but C2R continues
        if return_code != 0:
            c2r_procs = _find_running_clicktorun_processes()
            if c2r_procs:
                # Warn about setup exit but continue monitoring
                log.warning(
                    f"setup.exe exited with code {return_code}, but ClickToRun "
                    f"processes still running: {', '.join(p[1] for p in c2r_procs)}. "
                    "Continuing to monitor installation progress..."
                )

                # Update spinner to show we're monitoring C2R
                if _spinner is not None:
                    _spinner.update_task(
                        f"ODT: setup exited ({return_code}), monitoring ClickToRun..."
                    )

                # Stop the setup-based monitor and start C2R monitor
                stop_event.set()
                if monitor_thread and monitor_thread.is_alive():
                    monitor_thread.join(timeout=1.0)

                # Start new monitor for ClickToRun processes
                c2r_stop_event = threading.Event()
                if show_progress and (_spinner is not None):
                    c2r_monitor_thread = threading.Thread(
                        target=_monitor_clicktorun_progress,
                        args=(c2r_stop_event, product_names),
                        daemon=True,
                        name="c2r-progress",
                    )
                    c2r_monitor_thread.start()

                # Wait for ClickToRun processes to finish
                remaining_timeout = None
                if deadline:
                    remaining_timeout = max(0, deadline - time.monotonic())

                c2r_finished = _wait_for_clicktorun_processes(
                    c2r_stop_event, timeout=remaining_timeout
                )

                # Stop C2R monitor
                c2r_stop_event.set()
                if c2r_monitor_thread and c2r_monitor_thread.is_alive():
                    c2r_monitor_thread.join(timeout=1.0)

                if c2r_finished:
                    # C2R processes finished - check if installation succeeded
                    # by looking for Office files/registry
                    if _check_registry_key_exists(
                        r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
                    ):
                        log.info("ClickToRun installation appears to have completed successfully")
                        return_code = 0  # Override error code
                        error = None
                    else:
                        error = f"setup.exe exited with code {return_code}"
                else:
                    if remaining_timeout is not None and remaining_timeout <= 0:
                        error = f"Installation timed out after {timeout} seconds"
                    else:
                        error = f"ClickToRun processes did not complete (setup exit code: {return_code})"

    except subprocess.TimeoutExpired:
        if proc is not None:
            proc.kill()
            proc.wait()  # Wait for process to actually terminate
        stdout, stderr = "", ""
        return_code = -1
        error = f"Installation timed out after {timeout} seconds"
    except Exception as e:
        if proc is not None:
            try:
                proc.kill()
                proc.wait()
            except Exception:
                pass
        return_code = -1
        stdout = ""
        stderr = str(e)
        error = str(e)
    finally:
        # Stop monitoring
        stop_event.set()
        if monitor_thread and monitor_thread.is_alive():
            monitor_thread.join(timeout=1.0)

        # Unregister process
        if (_spinner is not None) and proc is not None:
            try:
                _spinner.unregister_process(proc)
            except Exception:
                pass

        # Clear spinner task
        if spinner_started and (_spinner is not None):
            _spinner.clear_task()

    duration = time.monotonic() - start_time

    # Clean up temp file on success
    if return_code == 0:
        try:
            config_path.unlink()
        except OSError:
            pass

    return ODTResult(
        success=return_code == 0,
        return_code=return_code,
        command=command,
        config_path=config_path if return_code != 0 else None,
        stdout=stdout,
        stderr=stderr,
        duration=duration,
        error=error,
    )


def _monitor_odt_download_progress(
    proc: subprocess.Popen[str],
    stop_event: threading.Event,
    download_path: Path,
) -> None:
    """!
    @brief Background thread to monitor ODT download progress.
    @details Uses a dedicated poller thread for I/O, this thread only updates spinner.
    @param proc The ODT subprocess being monitored.
    @param stop_event Event to signal when monitoring should stop.
    @param download_path Path where files are being downloaded.
    """
    if _spinner is None:
        return

    # Shared stats object for non-blocking reads
    stats = _MonitorStats()
    poller_stop = threading.Event()

    # Start dedicated poller thread for all I/O operations
    pid = proc.pid
    poller = threading.Thread(
        target=_download_poller_thread,
        args=(pid, download_path, stats, poller_stop),
        daemon=True,
        name="odt-download-poller",
    )
    poller.start()

    try:
        # Update spinner with stats every 200ms (non-blocking reads only)
        while not stop_event.is_set() and proc.poll() is None:
            # Read cached stats (non-blocking, thread-safe)
            cpu, mem, dl_size, dl_files, log_status, log_pct = stats.get_download()

            # Format size
            if dl_size >= 1_000_000_000:
                size_str = f"{dl_size / 1_000_000_000:.1f}GB"
            elif dl_size >= 1_000_000:
                size_str = f"{dl_size / 1_000_000:.0f}MB"
            else:
                size_str = f"{dl_size / 1_000:.0f}KB"

            # Build status line
            parts = ["ODT: Downloading"]
            if log_pct is not None:
                parts.append(f"{log_pct}%")
            if log_status and log_status != "Starting...":
                # Truncate long status
                status_short = log_status[:30] + "..." if len(log_status) > 33 else log_status
                parts.append(status_short)

            # Add metrics
            metrics = []
            if dl_size > 0:
                metrics.append(size_str)
            if dl_files > 0:
                metrics.append(f"{dl_files} files")
            if cpu > 0:
                metrics.append(f"CPU {cpu:.0f}%")
            if mem > 0:
                metrics.append(f"RAM {mem:.0f}MB")

            if metrics:
                parts.append(f"[{', '.join(metrics)}]")

            # Update spinner text (non-blocking)
            _spinner.update_task(" ".join(parts))

            # Short sleep to keep responsive
            stop_event.wait(0.2)
    finally:
        # Stop the poller thread
        poller_stop.set()
        poller.join(timeout=1.0)


def run_odt_download(
    config: ODTConfig,
    download_path: str | Path,
    *,
    dry_run: bool = False,
    timeout: float | None = None,
    show_progress: bool = True,
) -> ODTResult:
    """!
    @brief Run ODT setup.exe to download Office installation files.
    @param config ODTConfig with products and settings.
    @param download_path Local path to store downloaded files.
    @param dry_run If True, only generate the XML and print the command.
    @param timeout Optional timeout in seconds.
    @param show_progress If True, display spinner with progress updates.
    @returns ODTResult with execution details.
    """
    log = logging_ext.get_human_logger()

    try:
        setup_path = get_odt_setup_path()
    except FileNotFoundError as e:
        return ODTResult(
            success=False,
            return_code=-1,
            command=[],
            config_path=None,
            stdout="",
            stderr=str(e),
            duration=0.0,
            error=str(e),
        )

    download_path = Path(download_path)
    download_path.mkdir(parents=True, exist_ok=True)

    # Generate download XML
    xml_content = build_download_xml(config, str(download_path))

    fd, temp_path = tempfile.mkstemp(suffix=".xml", prefix="odt_download_")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(xml_content)
    except Exception:
        os.close(fd)
        raise

    config_path = Path(temp_path)
    command = [str(setup_path), "/download", str(config_path)]

    log.info(f"Downloading Office with ODT: {' '.join(command)}")
    log.info(f"Download path: {download_path}")

    if dry_run:
        log.info("[DRY-RUN] Would execute: %s", " ".join(command))
        log.info(f"Configuration XML:\n{xml_content}")
        return ODTResult(
            success=True,
            return_code=0,
            command=command,
            config_path=config_path,
            stdout="",
            stderr="",
            duration=0.0,
        )

    # Start spinner if available and requested
    spinner_started = False
    if show_progress and (_spinner is not None):
        _spinner.start_spinner_thread()
        _spinner.set_task("ODT: Starting download...")
        spinner_started = True

    start_time = time.monotonic()
    stop_event = threading.Event()
    monitor_thread: threading.Thread | None = None
    proc: subprocess.Popen[str] | None = None

    try:
        # Start the ODT process
        # Don't capture stdout/stderr - ODT writes to its own logs
        # Using DEVNULL prevents pipe buffer blocking issues on Windows
        proc = subprocess.Popen(
            command,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            text=True,
        )

        # Register for cleanup on Ctrl+C
        if _spinner is not None:
            _spinner.register_process(proc)

        # Start progress monitoring thread
        if show_progress and (_spinner is not None):
            monitor_thread = threading.Thread(
                target=_monitor_odt_download_progress,
                args=(proc, stop_event, download_path),
                daemon=True,
                name="odt-download-progress",
            )
            monitor_thread.start()

        # Wait for process with polling to keep threads responsive
        return_code = None
        deadline = time.monotonic() + timeout if timeout else None
        while return_code is None:
            return_code = proc.poll()
            if return_code is None:
                if deadline and time.monotonic() > deadline:
                    raise subprocess.TimeoutExpired(command, timeout)
                time.sleep(0.1)  # Poll every 100ms

        stdout, stderr = "", ""
        error = None

    except subprocess.TimeoutExpired:
        if proc is not None:
            proc.kill()
            proc.wait()  # Wait for process to actually terminate
        stdout, stderr = "", ""
        return_code = -1
        error = f"Download timed out after {timeout} seconds"
    except Exception as e:
        if proc is not None:
            try:
                proc.kill()
                proc.wait()
            except Exception:
                pass
        return_code = -1
        stdout = ""
        stderr = str(e)
        error = str(e)
    finally:
        # Stop monitoring
        stop_event.set()
        if monitor_thread and monitor_thread.is_alive():
            monitor_thread.join(timeout=1.0)

        # Unregister process
        if (_spinner is not None) and proc is not None:
            try:
                _spinner.unregister_process(proc)
            except Exception:
                pass

        # Clear spinner task
        if spinner_started and (_spinner is not None):
            _spinner.clear_task()

    duration = time.monotonic() - start_time

    # Clean up temp file
    try:
        config_path.unlink()
    except OSError:
        pass

    return ODTResult(
        success=return_code == 0,
        return_code=return_code,
        command=command,
        config_path=None,
        stdout=stdout,
        stderr=stderr,
        duration=duration,
        error=error,
    )


def run_odt_remove(
    *,
    remove_all: bool = True,
    product_ids: Sequence[str] | None = None,
    remove_msi: bool = True,
    dry_run: bool = False,
    timeout: float | None = None,
    show_progress: bool = True,
) -> ODTResult:
    """!
    @brief Run ODT setup.exe to remove Office installations.
    @param remove_all Remove all Office products.
    @param product_ids Specific product IDs to remove (if not remove_all).
    @param remove_msi Also remove MSI-based installations.
    @param dry_run If True, only generate the XML and print the command.
    @param timeout Optional timeout in seconds.
    @param show_progress If True, display spinner with progress updates.
    @returns ODTResult with execution details.
    """
    log = logging_ext.get_human_logger()

    try:
        setup_path = get_odt_setup_path()
    except FileNotFoundError as e:
        return ODTResult(
            success=False,
            return_code=-1,
            command=[],
            config_path=None,
            stdout="",
            stderr=str(e),
            duration=0.0,
            error=str(e),
        )

    # Generate removal XML
    xml_content = build_removal_xml(
        remove_all=remove_all,
        product_ids=product_ids,
        remove_msi=remove_msi,
    )

    fd, temp_path = tempfile.mkstemp(suffix=".xml", prefix="odt_remove_")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(xml_content)
    except Exception:
        os.close(fd)
        raise

    config_path = Path(temp_path)
    command = [str(setup_path), "/configure", str(config_path)]

    log.info(f"Removing Office with ODT: {' '.join(command)}")

    if dry_run:
        log.info("[DRY-RUN] Would execute: %s", " ".join(command))
        log.info(f"Configuration XML:\n{xml_content}")
        return ODTResult(
            success=True,
            return_code=0,
            command=command,
            config_path=config_path,
            stdout="",
            stderr="",
            duration=0.0,
        )

    # Start spinner if available and requested
    spinner_started = False
    if show_progress and (_spinner is not None):
        _spinner.start_spinner_thread()
        _spinner.set_task("ODT: Removing Office...")
        spinner_started = True

    start_time = time.monotonic()
    stop_event = threading.Event()
    monitor_thread: threading.Thread | None = None
    proc: subprocess.Popen[str] | None = None

    try:
        # Start the ODT process
        proc = subprocess.Popen(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )

        # Register for cleanup on Ctrl+C
        if _spinner is not None:
            _spinner.register_process(proc)

        # Start progress monitoring thread (reuse install monitor)
        if show_progress and (_spinner is not None):
            monitor_thread = threading.Thread(
                target=_monitor_odt_progress,
                args=(proc, stop_event, ["Removing Office"]),
                daemon=True,
                name="odt-remove-progress",
            )
            monitor_thread.start()

        # Wait for process to complete
        stdout, stderr = proc.communicate(timeout=timeout)
        return_code = proc.returncode
        error = None

    except subprocess.TimeoutExpired:
        if proc is not None:
            proc.kill()
            stdout, stderr = proc.communicate()
        else:
            stdout, stderr = "", ""
        return_code = -1
        error = f"Removal timed out after {timeout} seconds"
    except Exception as e:
        if proc is not None:
            try:
                proc.kill()
                proc.communicate()
            except Exception:
                pass
        return_code = -1
        stdout = ""
        stderr = str(e)
        error = str(e)
    finally:
        # Stop monitoring
        stop_event.set()
        if monitor_thread and monitor_thread.is_alive():
            monitor_thread.join(timeout=1.0)

        # Unregister process
        if (_spinner is not None) and proc is not None:
            try:
                _spinner.unregister_process(proc)
            except Exception:
                pass

        # Clear spinner task
        if spinner_started and (_spinner is not None):
            _spinner.clear_task()

    duration = time.monotonic() - start_time

    # Clean up temp file
    try:
        config_path.unlink()
    except OSError:
        pass

    return ODTResult(
        success=return_code == 0,
        return_code=return_code,
        command=command,
        config_path=None,
        stdout=stdout,
        stderr=stderr,
        duration=duration,
        error=error,
    )


def install_from_preset(
    preset_name: str,
    languages: list[str] | None = None,
    *,
    dry_run: bool = False,
    timeout: float | None = None,
) -> ODTResult:
    """!
    @brief Quick install using a preset configuration.
    @param preset_name Name of the preset from INSTALL_PRESETS.
    @param languages List of language codes (default: ["en-us"]).
    @param dry_run If True, only print what would be done.
    @param timeout Optional timeout in seconds.
    @returns ODTResult with execution details.
    """
    config = ODTConfig.from_preset(preset_name, languages)
    return run_odt_install(config, dry_run=dry_run, timeout=timeout)


def install_ltsc_2024_full(
    languages: list[str] | None = None,
    architecture: Architecture = Architecture.X64,
    *,
    dry_run: bool = False,
    timeout: float | None = None,
) -> ODTResult:
    """!
    @brief Quick install for Office LTSC 2024 + Visio + Project.
    @param languages List of language codes (default: ["en-us"]).
    @param architecture x64 or x86 (default: x64).
    @param dry_run If True, only print what would be done.
    @param timeout Optional timeout in seconds.
    @returns ODTResult with execution details.
    """
    langs = languages or ["en-us"]
    config = build_office_ltsc(
        "2024",
        architecture=architecture,
        languages=langs,
        volume=True,
        include_visio=True,
        include_project=True,
    )
    return run_odt_install(config, dry_run=dry_run, timeout=timeout)


# ---------------------------------------------------------------------------
# Listing and Discovery
# ---------------------------------------------------------------------------


def list_products() -> list[dict[str, object]]:
    """!
    @brief List all available product IDs with metadata.
    @returns List of product dictionaries.
    """
    result: list[dict[str, object]] = []
    for pid, meta in PRODUCT_IDS.items():
        result.append(
            {
                "id": pid,
                "name": meta.get("name", pid),
                "description": meta.get("description", ""),
                "channels": [ch.value for ch in meta.get("channels", [])],  # type: ignore[union-attr]
            }
        )
    return result


def list_presets() -> list[dict[str, object]]:
    """!
    @brief List all available installation presets.
    @returns List of preset dictionaries.
    """
    result: list[dict[str, object]] = []
    for name, preset in INSTALL_PRESETS.items():
        result.append(
            {
                "name": name,
                "products": preset.get("products", []),
                "architecture": preset.get("architecture", Architecture.X64).value,  # type: ignore[union-attr]
                "channel": preset.get("channel", UpdateChannel.CURRENT).value,  # type: ignore[union-attr]
                "description": preset.get("description", ""),
            }
        )
    return result


def list_channels() -> list[dict[str, str]]:
    """!
    @brief List all available update channels.
    @returns List of channel dictionaries.
    """
    return [{"name": ch.name, "value": ch.value} for ch in UpdateChannel]


def list_languages() -> list[str]:
    """!
    @brief List all supported language codes.
    @returns List of language code strings.
    """
    return list(SUPPORTED_LANGUAGES)


# ---------------------------------------------------------------------------
# Quick Builders
# ---------------------------------------------------------------------------


def build_365_proplus(
    *,
    architecture: Architecture = Architecture.X64,
    languages: list[str] | None = None,
    include_visio: bool = False,
    include_project: bool = False,
    shared_computer: bool = False,
    exclude_apps: list[str] | None = None,
) -> ODTConfig:
    """!
    @brief Quick builder for Microsoft 365 Apps for enterprise.
    @param architecture Target architecture (x64 or x86).
    @param languages Language codes (default: ["en-us"]).
    @param include_visio Include Visio Professional.
    @param include_project Include Project Professional.
    @param shared_computer Enable shared computer licensing.
    @param exclude_apps Apps to exclude (e.g., ["OneDrive", "Teams"]).
    @returns Configured ODTConfig instance.
    """
    langs = languages or ["en-us"]
    products = [
        ProductConfig("O365ProPlusRetail", languages=langs, exclude_apps=exclude_apps or [])
    ]

    if include_visio:
        products.append(ProductConfig("VisioProRetail", languages=langs))
    if include_project:
        products.append(ProductConfig("ProjectProRetail", languages=langs))

    return ODTConfig(
        products=products,
        architecture=architecture,
        channel=UpdateChannel.CURRENT,
        shared_computer_licensing=shared_computer,
    )


def build_office_ltsc(
    version: str,
    *,
    architecture: Architecture = Architecture.X64,
    languages: list[str] | None = None,
    volume: bool = True,
    include_visio: bool = False,
    include_project: bool = False,
) -> ODTConfig:
    """!
    @brief Quick builder for Office LTSC perpetual versions.
    @param version Office version: "2024", "2021", or "2019".
    @param architecture Target architecture.
    @param languages Language codes.
    @param volume Use volume license product (vs retail).
    @param include_visio Include Visio Professional.
    @param include_project Include Project Professional.
    @returns Configured ODTConfig instance.
    @raises ValueError for unsupported versions.
    """
    version_map = {
        "2024": (UpdateChannel.PERPETUAL_VL_2024, "ProPlus2024", "VisioPro2024", "ProjectPro2024"),
        "2021": (UpdateChannel.PERPETUAL_VL_2021, "ProPlus2021", "VisioPro2021", "ProjectPro2021"),
        "2019": (UpdateChannel.PERPETUAL_VL_2019, "ProPlus2019", "VisioPro2019", "ProjectPro2019"),
    }

    if version not in version_map:
        raise ValueError(f"Unsupported LTSC version: {version}. Use 2024, 2021, or 2019.")

    channel, office_base, visio_base, project_base = version_map[version]
    suffix = "Volume" if volume else "Retail"

    langs = languages or ["en-us"]
    products = [ProductConfig(f"{office_base}{suffix}", languages=langs)]

    if include_visio:
        products.append(ProductConfig(f"{visio_base}{suffix}", languages=langs))
    if include_project:
        products.append(ProductConfig(f"{project_base}{suffix}", languages=langs))

    return ODTConfig(
        products=products,
        architecture=architecture,
        channel=channel,
    )


# ---------------------------------------------------------------------------
# Module Entry Point
# ---------------------------------------------------------------------------


def _print_products() -> None:
    """Print all available products."""
    print("\nAvailable Office Products:")
    print("-" * 80)
    for product in list_products():
        channels = ", ".join(product.get("channels", []))  # type: ignore[arg-type]
        print(f"  {product['id']:<30} {product['name']}")
        print(f"      Channels: {channels}")
        print(f"      {product['description']}")
        print()


def _print_presets() -> None:
    """Print all available presets."""
    print("\nAvailable Installation Presets:")
    print("-" * 80)
    for preset in list_presets():
        products = ", ".join(preset.get("products", []))  # type: ignore[arg-type]
        print(f"  {preset['name']:<30}")
        print(f"      Products: {products}")
        print(f"      Architecture: {preset['architecture']}, Channel: {preset['channel']}")
        print(f"      {preset['description']}")
        print()


def _print_channels() -> None:
    """Print all available channels."""
    print("\nAvailable Update Channels:")
    print("-" * 60)
    for ch in list_channels():
        print(f"  {ch['name']:<25} {ch['value']}")


def _print_languages() -> None:
    """Print all supported languages."""
    print("\nSupported Languages:")
    print("-" * 40)
    langs = list_languages()
    for i in range(0, len(langs), 4):
        row = langs[i : i + 4]
        print("  " + "  ".join(f"{lang:<12}" for lang in row))


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        prog="odt_build",
        description="Generate Office Deployment Tool XML configurations.",
    )
    subparsers = parser.add_subparsers(dest="command", help="Available commands")

    # List command
    list_parser = subparsers.add_parser("list", help="List available options")
    list_parser.add_argument(
        "what",
        choices=["products", "presets", "channels", "languages", "all"],
        help="What to list",
    )

    # Build command
    build_parser = subparsers.add_parser("build", help="Build ODT XML configuration")
    build_parser.add_argument(
        "--preset",
        "-p",
        metavar="NAME",
        help="Use a preset configuration",
    )
    build_parser.add_argument(
        "--product",
        "-P",
        metavar="ID",
        action="append",
        dest="products",
        help="Product ID to install (can be repeated)",
    )
    build_parser.add_argument(
        "--language",
        "-l",
        metavar="CODE",
        action="append",
        dest="languages",
        help="Language code (can be repeated, default: en-us)",
    )
    build_parser.add_argument(
        "--arch",
        "-a",
        choices=["32", "64"],
        default="64",
        help="Architecture (32 or 64, default: 64)",
    )
    build_parser.add_argument(
        "--channel",
        "-c",
        metavar="NAME",
        help="Update channel (use 'list channels' to see options)",
    )
    build_parser.add_argument(
        "--output",
        "-o",
        metavar="FILE",
        help="Output file path (default: stdout)",
    )
    build_parser.add_argument(
        "--shared-computer",
        action="store_true",
        help="Enable shared computer licensing",
    )
    build_parser.add_argument(
        "--remove-msi",
        action="store_true",
        help="Remove existing MSI installations",
    )

    args = parser.parse_args()

    if args.command == "list":
        if args.what == "products" or args.what == "all":
            _print_products()
        if args.what == "presets" or args.what == "all":
            _print_presets()
        if args.what == "channels" or args.what == "all":
            _print_channels()
        if args.what == "languages" or args.what == "all":
            _print_languages()

    elif args.command == "build":
        try:
            if args.preset:
                config = ODTConfig.from_preset(args.preset, args.languages)
            elif args.products:
                langs = args.languages or ["en-us"]
                products = [ProductConfig(pid, languages=langs) for pid in args.products]
                arch = Architecture.X64 if args.arch == "64" else Architecture.X86

                channel = UpdateChannel.CURRENT
                if args.channel:
                    try:
                        channel = UpdateChannel[args.channel.upper().replace("-", "_")]
                    except KeyError:
                        for ch in UpdateChannel:
                            if ch.value.lower() == args.channel.lower():
                                channel = ch
                                break

                config = ODTConfig(
                    products=products,
                    architecture=arch,
                    channel=channel,
                    shared_computer_licensing=args.shared_computer,
                    remove_msi=args.remove_msi,
                )
            else:
                parser.error("Either --preset or --product must be specified")

            xml_output = build_xml(config)

            if args.output:
                Path(args.output).write_text(xml_output, encoding="utf-8")
                print(f"Configuration written to: {args.output}")
            else:
                print(xml_output)

        except ValueError as e:
            print(f"Error: {e}", file=sys.stderr)
            sys.exit(1)

    else:
        parser.print_help()
