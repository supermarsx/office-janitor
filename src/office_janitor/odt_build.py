"""!
@brief Office Deployment Tool XML configuration builder.
@details Generates ODT XML configuration files for installing Microsoft Office
products. Supports all product IDs and update channels available in the official
OEM configurations, allowing users to create custom installation configurations
programmatically.

@see https://docs.microsoft.com/en-us/deployoffice/office-deployment-tool-configuration-options
"""

from __future__ import annotations

import os
import re
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import TYPE_CHECKING

from . import constants, exec_utils, logging_ext

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


def run_odt_install(
    config: ODTConfig,
    *,
    dry_run: bool = False,
    timeout: float | None = None,
) -> ODTResult:
    """!
    @brief Run ODT setup.exe to install Office with the given configuration.
    @param config ODTConfig with products and settings.
    @param dry_run If True, only generate the XML and print the command.
    @param timeout Optional timeout in seconds.
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

    log.info(f"Installing Office with ODT: {' '.join(command)}")
    log.info(f"Config file: {config_path}")

    if dry_run:
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

    result = exec_utils.run_command(
        command,
        event="odt_install",
        timeout=timeout,
        human_message=f"Running ODT install: {config.products[0].product_id if config.products else 'unknown'}",
    )

    # Clean up temp file on success
    if result.returncode == 0:
        try:
            config_path.unlink()
        except OSError:
            pass

    return ODTResult(
        success=result.returncode == 0,
        return_code=result.returncode,
        command=command,
        config_path=config_path if result.returncode != 0 else None,
        stdout=result.stdout,
        stderr=result.stderr,
        duration=result.duration,
        error=result.error,
    )


def run_odt_download(
    config: ODTConfig,
    download_path: str | Path,
    *,
    dry_run: bool = False,
    timeout: float | None = None,
) -> ODTResult:
    """!
    @brief Run ODT setup.exe to download Office installation files.
    @param config ODTConfig with products and settings.
    @param download_path Local path to store downloaded files.
    @param dry_run If True, only generate the XML and print the command.
    @param timeout Optional timeout in seconds.
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

    result = exec_utils.run_command(
        command,
        event="odt_download",
        timeout=timeout,
        human_message=f"Downloading Office files to {download_path}",
    )

    # Clean up temp file
    try:
        config_path.unlink()
    except OSError:
        pass

    return ODTResult(
        success=result.returncode == 0,
        return_code=result.returncode,
        command=command,
        config_path=None,
        stdout=result.stdout,
        stderr=result.stderr,
        duration=result.duration,
        error=result.error,
    )


def run_odt_remove(
    *,
    remove_all: bool = True,
    product_ids: Sequence[str] | None = None,
    remove_msi: bool = True,
    dry_run: bool = False,
    timeout: float | None = None,
) -> ODTResult:
    """!
    @brief Run ODT setup.exe to remove Office installations.
    @param remove_all Remove all Office products.
    @param product_ids Specific product IDs to remove (if not remove_all).
    @param remove_msi Also remove MSI-based installations.
    @param dry_run If True, only generate the XML and print the command.
    @param timeout Optional timeout in seconds.
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

    result = exec_utils.run_command(
        command,
        event="odt_remove",
        timeout=timeout,
        human_message="Removing Office installations",
    )

    # Clean up temp file
    try:
        config_path.unlink()
    except OSError:
        pass

    return ODTResult(
        success=result.returncode == 0,
        return_code=result.returncode,
        command=command,
        config_path=None,
        stdout=result.stdout,
        stderr=result.stderr,
        duration=result.duration,
        error=result.error,
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
