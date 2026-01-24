# Office Click-to-Run Repair Guide

This document provides comprehensive documentation for the Office repair functionality in Office Janitor.

## Overview

Office Janitor supports automated repair of Microsoft Office Click-to-Run (C2R) installations. This feature uses the built-in `OfficeClickToRun.exe` utility and Office Deployment Tool (ODT) configurations to repair Office installations without requiring manual intervention through the Control Panel.

## Repair Types

### Quick Repair

**Command:** `office-janitor repair --quick`

Quick Repair is the recommended first-line approach for most Office issues. It:

- Runs locally without requiring internet connectivity
- Verifies and repairs local Office files
- Preserves user settings and customizations
- Does **NOT** reinstall excluded applications
- Typically completes in 5-15 minutes

**Best for:**
- Corrupted Office files
- Missing components
- Start-up issues
- Minor functionality problems

### Full Online Repair

**Command:** `office-janitor repair --full`

Full Online Repair is a more thorough repair option that reinstalls Office from Microsoft's CDN. It:

- Requires active internet connection
- Downloads and reinstalls Office components (~2-4 GB)
- **WARNING:** May reinstall previously excluded applications
- Takes 30-60 minutes depending on connection speed
- Resets many Office settings to defaults

**Best for:**
- Persistent issues not resolved by Quick Repair
- Severe corruption
- Missing Office features
- COM add-in registration issues

## Command-Line Options

### Basic Repair Commands

```powershell
# Quick Repair (recommended first)
office-janitor repair --quick

# Full Online Repair
office-janitor repair --full

# Dry-run mode (preview without executing)
office-janitor repair --quick --dry-run
```

### Advanced Options

```powershell
# Specify language/culture
office-janitor repair --quick --culture de-de

# Force specific architecture
office-janitor repair --quick --platform x64

# Show repair UI instead of silent mode
office-janitor repair --full --visible

# Use custom XML configuration
office-janitor repair --config-file "C:\path\to\config.xml"

# Skip confirmation prompts
office-janitor repair --quick --yes
```

### Option Reference

| Option | Description | Default |
|--------|-------------|---------|
| `--quick` | Quick local repair | - |
| `--full` | Full online repair from CDN | - |
| `--config-file XML` | Custom XML configuration file | - |
| `--culture LANG` | Language code (e.g., `en-us`, `de-de`) | `en-us` |
| `--platform ARCH` | Architecture: `x86` or `x64` | Auto-detected |
| `--visible` | Show repair progress UI | Silent |
| `--dry-run` | Simulate without executing | Disabled |
| `--yes` | Skip confirmation prompts | Disabled |

## Supported Languages/Cultures

The following language codes are supported for the `--culture` option:

| Code | Language |
|------|----------|
| `ar-sa` | Arabic (Saudi Arabia) |
| `bg-bg` | Bulgarian |
| `zh-cn` | Chinese (Simplified) |
| `zh-tw` | Chinese (Traditional) |
| `hr-hr` | Croatian |
| `cs-cz` | Czech |
| `da-dk` | Danish |
| `nl-nl` | Dutch |
| `en-us` | English (US) |
| `en-gb` | English (UK) |
| `et-ee` | Estonian |
| `fi-fi` | Finnish |
| `fr-fr` | French |
| `de-de` | German |
| `el-gr` | Greek |
| `he-il` | Hebrew |
| `hi-in` | Hindi |
| `hu-hu` | Hungarian |
| `id-id` | Indonesian |
| `it-it` | Italian |
| `ja-jp` | Japanese |
| `kk-kz` | Kazakh |
| `ko-kr` | Korean |
| `lv-lv` | Latvian |
| `lt-lt` | Lithuanian |
| `ms-my` | Malay |
| `nb-no` | Norwegian (Bokmål) |
| `pl-pl` | Polish |
| `pt-br` | Portuguese (Brazil) |
| `pt-pt` | Portuguese (Portugal) |
| `ro-ro` | Romanian |
| `ru-ru` | Russian |
| `sr-latn-rs` | Serbian (Latin) |
| `sk-sk` | Slovak |
| `sl-si` | Slovenian |
| `es-es` | Spanish (Spain) |
| `sv-se` | Swedish |
| `th-th` | Thai |
| `tr-tr` | Turkish |
| `uk-ua` | Ukrainian |
| `vi-vn` | Vietnamese |

## XML Configuration Files

### Bundled Configurations

Office Janitor includes pre-built XML configurations in the `oem/` folder:

| File | Description |
|------|-------------|
| `FullRemoval.xml` | Complete Office removal |
| `QuickRepair.xml` | Standard quick repair configuration |
| `FullRepair.xml` | Full online repair configuration |
| `Repair_ProPlus_x64.xml` | Office 365 ProPlus 64-bit |
| `Repair_ProPlus_x86.xml` | Office 365 ProPlus 32-bit |
| `Repair_ProPlus_Visio_Project.xml` | ProPlus with Visio and Project |
| `Repair_Business_x64.xml` | Microsoft 365 Business |
| `Repair_Office2019_x64.xml` | Office 2019 LTSC |
| `Repair_Office2021_x64.xml` | Office 2021 LTSC |
| `Repair_Office2024_x64.xml` | Office 2024 LTSC |
| `Repair_Multilang.xml` | Multi-language installation |
| `Repair_SharedComputer.xml` | Shared Computer Licensing |
| `Repair_Interactive.xml` | Full UI mode for troubleshooting |

### OEM Configuration CLI Aliases

For convenience, bundled configurations can be executed directly via CLI aliases:

```bash
# Execute OEM config by preset name
office-janitor --oem-config full-removal
office-janitor --oem-config quick-repair
office-janitor --oem-config proplus-x64

# Shortcut aliases for common operations
office-janitor --c2r-remove          # Complete C2R Office removal
office-janitor --c2r-repair-quick    # Quick repair
office-janitor --c2r-repair-full     # Full online repair
office-janitor --c2r-proplus         # Repair ProPlus x64
office-janitor --c2r-business        # Repair Business x64

# With dry-run to preview
office-janitor --oem-config full-removal --dry-run
```

#### Available Preset Names

| Preset Name | Config File | Description |
|-------------|-------------|-------------|
| `full-removal` | `FullRemoval.xml` | Removes all C2R Office |
| `quick-repair` | `QuickRepair.xml` | Local quick repair |
| `full-repair` | `FullRepair.xml` | Online full repair |
| `proplus-x64` | `Repair_ProPlus_x64.xml` | ProPlus 64-bit repair |
| `proplus-x86` | `Repair_ProPlus_x86.xml` | ProPlus 32-bit repair |
| `proplus-visio-project` | `Repair_ProPlus_Visio_Project.xml` | Suite with Visio/Project |
| `business-x64` | `Repair_Business_x64.xml` | M365 Business repair |
| `office2019-x64` | `Repair_Office2019_x64.xml` | Office 2019 LTSC |
| `office2021-x64` | `Repair_Office2021_x64.xml` | Office 2021 LTSC |
| `office2024-x64` | `Repair_Office2024_x64.xml` | Office 2024 LTSC |
| `multilang` | `Repair_Multilang.xml` | Multi-language support |
| `shared-computer` | `Repair_SharedComputer.xml` | Shared licensing |
| `interactive` | `Repair_Interactive.xml` | Full UI mode |

### Custom XML Configuration

You can create custom XML configurations for specific scenarios. Example structure:

```xml
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365ProPlusRetail">
      <Language ID="en-us" />
    </Product>
  </Add>
  
  <Updates Enabled="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Logging Level="Standard" Path="%temp%" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>
```

### Product IDs Reference

| Product | Product ID |
|---------|------------|
| Microsoft 365 Apps for Enterprise | `O365ProPlusRetail` |
| Microsoft 365 Apps for Business | `O365BusinessRetail` |
| Office 2021 ProPlus (Volume) | `ProPlus2021Volume` |
| Office 2019 ProPlus (Volume) | `ProPlus2019Volume` |
| Office 2024 ProPlus (Volume) | `ProPlus2024Volume` |
| Visio Professional | `VisioProRetail` |
| Project Professional | `ProjectProRetail` |
| Visio 2021 (Volume) | `VisioPro2021Volume` |
| Project 2021 (Volume) | `ProjectPro2021Volume` |

### Update Channels

| Channel Name | XML Value |
|--------------|-----------|
| Current Channel | `Current` |
| Monthly Enterprise | `MonthlyEnterprise` |
| Semi-Annual Enterprise | `SemiAnnual` |
| Office 2019 LTSC | `PerpetualVL2019` |
| Office 2021 LTSC | `PerpetualVL2021` |
| Office 2024 LTSC | `PerpetualVL2024` |

## Programmatic API

### Basic Usage

```python
from office_janitor import repair

# Quick repair
result = repair.quick_repair()
if result.success:
    print("Repair completed successfully")

# Full repair
result = repair.full_repair(culture="en-us", silent=True)

# Check result
print(result.summary)
print(f"Return code: {result.return_code}")
print(f"Duration: {result.duration}s")
```

### Advanced Configuration

```python
from office_janitor.repair import RepairConfig, RepairType, run_repair

# Create custom configuration
config = RepairConfig(
    repair_type=RepairType.QUICK,
    platform="x64",
    culture="en-us",
    force_app_shutdown=True,
)

# Execute repair
result = run_repair(config, dry_run=False)
```

### Using ODT Setup.exe

```python
from pathlib import Path
from office_janitor import repair

# Reconfigure using XML
config_path = Path("oem/Repair_ProPlus_x64.xml")
result = repair.reconfigure_office(config_path)
```

### Detection Functions

```python
from office_janitor import repair

# Check if C2R Office is installed
if repair.is_c2r_office_installed():
    print("C2R Office detected")

# Get installation details
info = repair.get_installed_c2r_info()
print(f"Version: {info['version']}")
print(f"Platform: {info['platform']}")
print(f"Culture: {info['culture']}")
print(f"Products: {info['product_ids']}")
```

## Direct Command-Line Usage

### Using OfficeClickToRun.exe

The repair functionality is based on `OfficeClickToRun.exe`, which can be invoked directly:

```powershell
# Location
$exe = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"

# Quick Repair (silent)
& $exe scenario=Repair platform=x64 culture=en-us RepairType=QuickRepair forceappshutdown=True DisplayLevel=False

# Full Repair (with UI)
& $exe scenario=Repair platform=x64 culture=en-us RepairType=FullRepair forceappshutdown=True DisplayLevel=True
```

### Parameters

| Parameter | Values | Required |
|-----------|--------|----------|
| `scenario` | `Repair` | Yes |
| `platform` | `x86`, `x64` | Yes |
| `culture` | Language code (e.g., `en-us`) | Yes |
| `RepairType` | `QuickRepair`, `FullRepair` | No (defaults to Quick) |
| `forceappshutdown` | `True`, `False` | No |
| `DisplayLevel` | `True`, `False` | No |

### Using Setup.exe (ODT)

```powershell
# Location (bundled or system)
$setup = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\setup.exe"

# Reconfigure/Repair
& $setup /configure "C:\path\to\config.xml"
```

## Troubleshooting

### Common Issues

#### "No Click-to-Run Office installation found"

This error occurs when Office Janitor cannot detect a C2R installation. This happens with:
- MSI-based Office installations (Office 2016 and earlier perpetual licenses)
- Office installed via Windows Store
- No Office installation present

**Solution:** For MSI installations, use the traditional repair:
1. Open Control Panel > Programs > Programs and Features
2. Select Microsoft Office
3. Click Change > Repair

#### "OfficeClickToRun.exe not found"

The repair executable was not found in standard locations.

**Solutions:**
1. Verify Office is installed correctly
2. Check `C:\Program Files\Common Files\Microsoft Shared\ClickToRun\`
3. Use the bundled executable in `oem/` folder

#### Repair hangs or times out

**Solutions:**
1. Try running with `--repair-visible` to see progress
2. Ensure no Office applications are running
3. Check Windows Update is not actively running
4. Increase timeout with `--timeout 7200`

#### Full Repair reinstalled excluded apps

This is expected behavior. Full Online Repair downloads a complete Office installation.

**Solutions:**
1. After repair, run setup.exe with a configuration that excludes unwanted apps
2. Use Quick Repair when possible
3. Create a custom XML configuration that maintains your exclusions

### Logging

Repair operations are logged to:
- Human-readable log: `%LOCALAPPDATA%\OfficeJanitor\Logs\office-janitor.log`
- Machine-readable JSONL: `%LOCALAPPDATA%\OfficeJanitor\Logs\office-janitor.jsonl`

Enable verbose logging:
```powershell
office-janitor --repair quick --json
```

### Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Success |
| 1 | Repair failed |
| 130 | Cancelled by user (Ctrl+C) |

## Best Practices

1. **Try Quick Repair first** - It's faster and preserves your settings
2. **Back up important files** - Although rare, repairs can occasionally affect user data
3. **Close all Office applications** - Repairs will force-close them anyway
4. **Run as Administrator** - Repair operations require elevated privileges
5. **Use dry-run mode first** - Preview what will happen: `--dry-run`
6. **Check connectivity** - Full Repair requires stable internet

## References

- [Microsoft Office Deployment Tool Documentation](https://docs.microsoft.com/en-us/deployoffice/overview-office-deployment-tool)
- [Office 365 Click-to-Run Deployment Guide](https://itpro.outsidesys.com/2016/05/18/deploying-office-365-click-to-run/)
- [Office Configuration XML Reference](https://docs.microsoft.com/en-us/deployoffice/office-deployment-tool-configuration-options)
- [Office Repair Options](https://support.microsoft.com/en-us/office/repair-an-office-application-7821d4b6-7c1d-4205-aa0e-a6b40c5bb88b)
