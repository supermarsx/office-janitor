# Office Janitor

[![CI](https://img.shields.io/github/actions/workflow/status/supermarsx/office-janitor/ci.yml?style=flat-square&label=CI)](https://github.com/supermarsx/office-janitor/actions/workflows/ci.yml)
[![GitHub stars](https://img.shields.io/github/stars/supermarsx/office-janitor?style=flat-square)](https://github.com/supermarsx/office-janitor/stargazers)
[![GitHub forks](https://img.shields.io/github/forks/supermarsx/office-janitor?style=flat-square)](https://github.com/supermarsx/office-janitor/network/members)
[![GitHub watchers](https://img.shields.io/github/watchers/supermarsx/office-janitor?style=flat-square)](https://github.com/supermarsx/office-janitor/watchers)
[![GitHub Downloads](https://img.shields.io/github/downloads/supermarsx/office-janitor/total?style=flat-square)](https://github.com/supermarsx/office-janitor/releases)
[![PyPI Downloads](https://img.shields.io/pepy/dt/office-janitor?style=flat-square&label=PyPI%20Downloads)](https://pepy.tech/project/office-janitor)
[![GitHub Issues](https://img.shields.io/github/issues/supermarsx/office-janitor?style=flat-square)](https://github.com/supermarsx/office-janitor/issues)
[![License](https://img.shields.io/github/license/supermarsx/office-janitor?style=flat-square)](license.md)

<img width="1017" height="902" alt="image" src="https://github.com/user-attachments/assets/37748fbb-f3a6-446b-81ec-1c2780e7137b" />

<br>
<br>

**Office Janitor** is a comprehensive, stdlib-only Python utility for managing Microsoft Office installations on Windows. It provides three core capabilities:

- **🔧 Install** – Deploy Office using ODT with presets, custom configurations, and live progress monitoring
- **🔄 Repair** – Quick and full repair for Click-to-Run installations with bundled OEM configurations
- **🧹 Scrub** – Deep uninstall and cleanup of MSI and Click-to-Run Office across all versions (2003-2024, Microsoft 365)

The tool follows the architecture defined in [`spec.md`](spec.md) and can be packaged into a single-file Windows executable with PyInstaller.

---

## Table of Contents

- [Installation](#installation)
- [Quick Start](#quick-start)
- [ODT Installation](#odt-installation)
  - [Installation Presets](#installation-presets)
  - [Quick Install Aliases](#quick-install-aliases)
  - [Custom Installations](#custom-installations)
  - [Progress Monitoring](#progress-monitoring)
- [Office Repair](#office-repair)
  - [Quick Repair](#quick-repair)
  - [Full Online Repair](#full-online-repair)
  - [OEM Configurations](#oem-configurations)
- [Office Removal & Scrubbing](#office-removal--scrubbing)
  - [Automatic Removal](#automatic-removal)
  - [Targeted Removal](#targeted-removal)
  - [Scrub Levels](#scrub-levels)
  - [License Management](#license-management)
- [CLI Reference](#cli-reference)
- [Configuration Files](#configuration-files)
- [Safety Guidance](#safety-guidance)
- [Logging & Diagnostics](#logging--diagnostics)
- [Contributing & Testing](#contributing--testing)

---

## Installation

### Prerequisites
- Windows 7 or later with administrator privileges
- Python 3.9+ (for running from source)
- Optional: PyInstaller for building standalone executables

### Running from Source
```bash
# Clone and enter the repository
git clone https://github.com/supermarsx/office-janitor.git
cd office-janitor

# Create virtual environment (recommended)
python -m venv .venv
.venv\Scripts\activate

# Install in editable mode
python -m pip install -e .

# Run the tool
python office_janitor.py --help
```

### Building Standalone Executable
```bash
pyinstaller --onefile --uac-admin --name office-janitor office_janitor.py --paths src
```
The resulting `dist/office-janitor.exe` is a single-file admin-elevated executable that includes the embedded ODT setup.exe.

---

## Quick Start

```bash
# Install Office LTSC 2024 with Visio and Project (no bloatware)
office-janitor install --goobler

# Repair Office Click-to-Run (quick repair)
office-janitor repair --quick

# Remove all Office installations (preview first!)
office-janitor remove --dry-run
office-janitor remove --backup C:\Backups

# Diagnose Office installations without making changes
office-janitor diagnose --plan report.json

# Interactive mode - launches menu
office-janitor
```

---

## ODT Installation

Office Janitor includes an embedded Office Deployment Tool (setup.exe) and can install any Office product directly with live progress monitoring.

### Installation Presets

Use presets for one-command installations:

```bash
# Install Office LTSC 2024 Professional Plus (64-bit)
office-janitor install --preset office2024-x64

# Install full suite: Office 2024 + Visio + Project
office-janitor install --preset ltsc2024-full-x64

# Install clean version without OneDrive/Skype bloatware
office-janitor install --preset ltsc2024-full-x64-clean

# Add multiple languages
office-janitor install --preset ltsc2024-full-x64 \
  --language en-us --language de-de --language es-mx

# Preview without installing
office-janitor install --preset office2024-x64 --dry-run
```

#### Microsoft 365 Presets

| Preset | Products | Description |
|--------|----------|-------------|
| `365-proplus-x64` | O365ProPlusRetail | Microsoft 365 Apps for enterprise (64-bit) |
| `365-proplus-x86` | O365ProPlusRetail | Microsoft 365 Apps for enterprise (32-bit) |
| `365-business-x64` | O365BusinessRetail | Microsoft 365 Apps for business |
| `365-proplus-visio-project` | O365ProPlusRetail + Visio + Project | Full M365 suite |
| `365-shared-computer` | O365ProPlusRetail | Shared Computer Licensing enabled |
| `365-proplus-x64-clean` | O365ProPlusRetail | **No OneDrive/Skype** |
| `365-proplus-visio-project-clean` | Full suite | **No OneDrive/Skype** |

#### Office LTSC 2024 Presets

| Preset | Products | Description |
|--------|----------|-------------|
| `office2024-x64` | ProPlus2024Volume | Office LTSC 2024 Professional Plus |
| `office2024-x86` | ProPlus2024Volume | 32-bit version |
| `office2024-standard-x64` | Standard2024Volume | Standard edition |
| `ltsc2024-full-x64` | ProPlus + Visio + Project | **Complete 2024 suite** |
| `ltsc2024-full-x86` | ProPlus + Visio + Project | 32-bit complete suite |
| `ltsc2024-x64-clean` | ProPlus2024Volume | **No OneDrive/Skype** |
| `ltsc2024-full-x64-clean` | Full suite | **No OneDrive/Skype** ⭐ |
| `ltsc2024-full-x86-clean` | Full suite (32-bit) | **No bloatware** |

#### Office LTSC 2021 & 2019 Presets

| Preset | Products | Description |
|--------|----------|-------------|
| `office2021-x64` | ProPlus2021Volume | Office LTSC 2021 Professional Plus |
| `office2021-standard-x64` | Standard2021Volume | Standard edition |
| `ltsc2021-full-x64` | ProPlus + Visio + Project | Complete 2021 suite |
| `office2019-x64` | ProPlus2019Volume | Office 2019 Professional Plus |
| `office2019-x86` | ProPlus2019Volume | 32-bit version |

#### Standalone Products

| Preset | Product | Description |
|--------|---------|-------------|
| `visio-pro-x64` | VisioPro2024Volume | Visio Professional 2024 |
| `project-pro-x64` | ProjectPro2024Volume | Project Professional 2024 |

### Quick Install Aliases

Author-defined shortcuts for common installations:

```bash
# Goobler: Full Office 2024 suite, no bloatware, Portuguese + English
office-janitor install --goobler

# Pupa: ProPlus only, no bloatware, Portuguese + English  
office-janitor install --pupa

# Both support dry-run
office-janitor install --goobler --dry-run
```

| Alias | Preset | Products | Languages |
|-------|--------|----------|-----------|
| `--goobler` | `ltsc2024-full-x64-clean` | ProPlus 2024 + Visio + Project | pt-pt, en-us |
| `--pupa` | `ltsc2024-x64-clean` | ProPlus 2024 only | pt-pt, en-us |

### Custom Installations

Build custom configurations when presets don't fit:

```bash
# Custom product selection
office-janitor install \
  --product ProPlus2024Volume \
  --product VisioPro2024Volume \
  --channel PerpetualVL2024 \
  --language en-us \
  --exclude-app OneDrive \
  --exclude-app Lync

# Generate XML without installing
office-janitor odt --build --preset office2024-x64 --output install.xml

# Download for offline installation
office-janitor odt --download "D:\OfficeSource" \
  --preset 365-proplus-x64 \
  --language en-us --language es-es

# Generate removal XML
office-janitor odt --removal --remove-msi --output remove.xml
```

### Progress Monitoring

During installation, Office Janitor provides real-time progress:

```
⠋ ODT: ProPlus2024Volume, VisioPro2024Volume 45% Installing Office... [1.2GB, 3421 files, 892 keys, CPU 12%, RAM 245MB] (5m 34s)
```

The spinner shows:
- **Products** being installed
- **Progress percentage** from ODT logs
- **Current phase** (downloading, installing, configuring)
- **Disk usage** (Office installation size)
- **File count** in Office directories
- **Registry keys** created
- **CPU/RAM** usage of installer processes
- **Elapsed time**

If setup.exe exits but ClickToRun processes continue (common behavior), monitoring automatically switches to track those processes until installation completes.

**Ctrl+C** during installation will:
1. Terminate the ODT setup process
2. Kill all ClickToRun-related processes (OfficeClickToRun.exe, OfficeC2RClient.exe, etc.)
3. Display what was terminated

---

## Office Repair

Repair Click-to-Run Office installations without reinstalling.

### Quick Repair

Fast local repair using cached installation files:

```bash
# Quick repair (runs silently)
office-janitor repair --quick

# Show repair UI
office-janitor repair --quick --visible

# Preview without executing
office-janitor repair --quick --dry-run
```

Quick repair:
- Uses locally cached files (fast, no download)
- Fixes corrupted files and settings
- Preserves user data and customizations
- Completes in 5-15 minutes

### Full Online Repair

Complete repair that re-downloads Office from CDN:

```bash
# Full online repair
office-janitor repair --full

# With visible progress UI
office-janitor repair --full --visible

# Specify architecture
office-janitor repair --full --platform x64
```

Full repair:
- Downloads fresh files from Microsoft CDN
- Repairs more severe corruption
- Takes 30-60+ minutes depending on connection
- Requires internet connectivity

### OEM Configurations

Use bundled configuration presets for repair/reconfiguration:

```bash
# List available OEM presets
office-janitor repair --help

# Quick repair preset
office-janitor repair --config quick-repair

# Full repair preset  
office-janitor repair --config full-repair

# Repair specific products
office-janitor repair --config proplus-x64
office-janitor repair --config business-x64
office-janitor repair --config office2024-x64

# Remove all C2R products
office-janitor c2r --remove

# Quick aliases
office-janitor repair --quick
office-janitor repair --full
```

Available OEM presets:
- `full-removal` - Remove all C2R Office products
- `quick-repair` - Quick local repair
- `full-repair` - Full online repair
- `proplus-x64` / `proplus-x86` - Repair Office 365 ProPlus
- `proplus-visio-project` - Repair full suite
- `business-x64` - Repair Microsoft 365 Business
- `office2019-x64` - Repair Office 2019
- `office2021-x64` - Repair Office 2021
- `office2024-x64` - Repair Office 2024
- `multilang` - Multi-language configuration
- `shared-computer` - Shared Computer Licensing
- `interactive` - Show repair UI

### Custom Repair Configuration

Use your own XML configuration:

```bash
office-janitor repair --config-file "C:\Configs\custom_repair.xml"
```

---

## Office Removal & Scrubbing

Deep uninstall and cleanup for all Office versions (2003-2024, Microsoft 365).

### Automatic Removal

Remove all detected Office installations:

```bash
# ALWAYS preview first!
office-janitor remove --dry-run

# Execute with backup
office-janitor remove --backup "C:\Backups\Office"

# Silent unattended removal
office-janitor remove --yes --quiet

# Keep user data during removal
office-janitor remove --keep-templates --keep-user-settings --keep-license
```

### Targeted Removal

Remove specific Office versions:

```bash
# Remove only Office 2016
office-janitor remove --target 2016

# Remove Microsoft 365 only
office-janitor remove --target 365

# Remove Office 2019 including Visio/Project
office-janitor remove --target 2019 --include visio,project

# Remove only MSI-based Office
office-janitor remove --msi-only

# Remove only Click-to-Run Office
office-janitor remove --c2r-only

# Remove specific MSI product by GUID
office-janitor remove --product-code "{90160000-0011-0000-0000-0000000FF1CE}"

# Remove specific C2R release
office-janitor remove --release-id O365ProPlusRetail
```

### Scrub Levels

Control cleanup intensity:

```bash
# Minimal - uninstall only
office-janitor remove --scrub-level minimal

# Standard - uninstall + residue cleanup (default)
office-janitor remove --scrub-level standard

# Aggressive - deep registry/filesystem cleanup
office-janitor remove --scrub-level aggressive

# Nuclear - remove everything possible
office-janitor remove --scrub-level nuclear
```

| Level | Uninstall | Files | Registry | Services | Tasks | Licenses |
|-------|-----------|-------|----------|----------|-------|----------|
| minimal | ✅ | ❌ | ❌ | ❌ | ❌ | ❌ |
| standard | ✅ | ✅ | ✅ | ✅ | ✅ | ❌ |
| aggressive | ✅ | ✅ | ✅+ | ✅ | ✅ | ✅ |
| nuclear | ✅ | ✅+ | ✅++ | ✅ | ✅ | ✅+ |

### Cleanup-Only Mode

Skip uninstall, clean residue only:

```bash
# Clean leftover files/registry after manual uninstall
office-janitor remove --cleanup-only

# Aggressive residue cleanup
office-janitor remove --cleanup-only --scrub-level aggressive

# Registry cleanup only
office-janitor remove --registry-only
```

### License Management

```bash
# Clean all Office licenses
office-janitor license --clean-all

# Clean SPP tokens only (KMS/MAK)
office-janitor license --clean-spp

# Clean OSPP tokens
office-janitor license --clean-ospp

# Clean vNext/device-based licensing
office-janitor license --clean-vnext

# Preserve licenses during removal
office-janitor remove --keep-license
```

### Additional Cleanup Options

```bash
# Clean MSOCache installation files
office-janitor remove --clean-msocache

# Remove Office AppX/MSIX packages
office-janitor remove --clean-appx

# Clean Windows Installer metadata
office-janitor remove --clean-wi-metadata

# Clean Office shortcuts
office-janitor remove --clean-shortcuts

# Clean registry add-ins, COM, shell extensions
office-janitor remove --cleanup-only --clean-addin-registry --clean-com-registry --clean-shell-extensions
```

### Multiple Passes

For stubborn installations:

```bash
# Run 3 uninstall passes
office-janitor remove --passes 3

# Or use max-passes
office-janitor remove --max-passes 5
```

---

## CLI Reference

Office Janitor uses a subcommand-based interface. Run `office-janitor <command> --help` for command-specific options.

### Commands

| Command | Description |
|---------|-------------|
| `install` | Install Office via ODT with presets or custom configurations |
| `repair` | Repair Click-to-Run Office installations |
| `remove` | Remove and scrub Office installations |
| `diagnose` | Detection and planning only, no changes |
| `odt` | Generate ODT XML configurations |
| `offscrub` | Legacy OffScrub compatibility mode |
| `c2r` | Click-to-Run management operations |
| `license` | Office license management |
| `config` | Manage configuration files |
| (none) | Launch interactive menu |

### Global Options

Available with all commands:

| Flag | Description |
|------|-------------|
| `-n, --dry-run` | Simulate without changes |
| `-y, --yes` | Skip confirmations |
| `--config JSON` | Load options from file |
| `--logdir DIR` | Custom log directory |
| `--timeout SEC` | Per-step timeout |
| `-v, -vv, -vvv` | Increase verbosity |
| `--quiet` | Reduce output |
| `--no-color` | Disable colors |

### Install Command

Install Office using Office Deployment Tool:

```bash
office-janitor install [OPTIONS]
```

| Flag | Description |
|------|-------------|
| `--preset NAME` | Use installation preset |
| `--product ID` | Add product (repeatable) |
| `--language CODE` | Add language (repeatable) |
| `--arch 32/64` | Architecture (default: 64) |
| `--channel CHANNEL` | Update channel |
| `--exclude-app APP` | Exclude app (repeatable) |
| `--shared-computer` | Enable shared licensing |
| `--goobler` | Full LTSC 2024 suite, no bloatware (pt-pt, en-us) |
| `--pupa` | ProPlus 2024 only, no bloatware (pt-pt, en-us) |
| `--list-presets` | List available presets |
| `--list-products` | List product IDs |
| `--list-channels` | List update channels |
| `--list-languages` | List language codes |

### Repair Command

Repair Click-to-Run Office installations:

```bash
office-janitor repair [OPTIONS]
```

| Flag | Description |
|------|-------------|
| `--quick` | Quick local repair |
| `--full` | Full online repair from CDN |
| `--config NAME` | Use OEM configuration preset |
| `--config-file XML` | Custom XML configuration |
| `--culture LANG` | Language for repair (default: en-us) |
| `--platform ARCH` | Architecture (x86/x64) |
| `--visible` | Show repair UI |
| `--timeout SEC` | Timeout (default: 3600) |

### Remove Command

Remove and scrub Office installations:

```bash
office-janitor remove [OPTIONS]
```

| Flag | Description |
|------|-------------|
| `--target VER` | Target specific version (2003-2024, 365) |
| `--msi-only` | Remove only MSI-based Office |
| `--c2r-only` | Remove only Click-to-Run Office |
| `--product-code GUID` | Remove specific MSI product |
| `--release-id ID` | Remove specific C2R release |
| `--scrub-level LEVEL` | minimal/standard/aggressive/nuclear |
| `--passes N` | Uninstall passes |
| `--backup DIR` | Backup registry/files |
| `--cleanup-only` | Skip uninstall, clean residue only |
| `--registry-only` | Only registry cleanup |
| `--skip-uninstall` | Skip uninstall phase |
| `--skip-processes` | Don't terminate Office processes |
| `--skip-services` | Don't stop Office services |
| `--skip-tasks` | Don't remove scheduled tasks |
| `--skip-registry` | Don't clean registry |
| `--skip-filesystem` | Don't clean files |

### Diagnose Command

Detection and planning without changes:

```bash
office-janitor diagnose [OPTIONS]
```

| Flag | Description |
|------|-------------|
| `--plan FILE` | Export plan to JSON |
| `--json` | JSON output to stdout |

### ODT Command

Generate ODT XML configurations:

```bash
office-janitor odt [OPTIONS]
```

| Flag | Description |
|------|-------------|
| `--build` | Generate configuration XML |
| `--download DIR` | Download Office source files |
| `--removal` | Generate removal XML |
| `--remove-msi` | Include RemoveMSI element |
| `--output FILE` | Output XML path |
| `--preset NAME` | Use preset configuration |
| `--product ID` | Add product (repeatable) |
| `--language CODE` | Add language (repeatable) |

### License Command

Office license management:

```bash
office-janitor license [OPTIONS]
```

| Flag | Description |
|------|-------------|
| `--clean-all` | Clean all license types |
| `--clean-spp` | Clean SPP tokens (KMS/MAK) |
| `--clean-ospp` | Clean OSPP tokens |
| `--clean-vnext` | Clean vNext/device licensing |
| `--keep-license` | Preserve licenses during removal |

### C2R Command

Click-to-Run management:

```bash
office-janitor c2r [OPTIONS]
```

| Flag | Description |
|------|-------------|
| `--remove` | Remove all C2R products |
| `--repair` | Repair C2R installation |

### User Data Options

Available with `remove` command:

| Flag | Description |
|------|-------------|
| `--keep-templates` | Preserve templates |
| `--keep-user-settings` | Preserve settings |
| `--keep-outlook-data` | Preserve Outlook data |
| `--delete-user-settings` | Remove settings |
| `--clean-shortcuts` | Remove shortcuts |

### Cleanup Options

Available with `remove` command:

| Flag | Description |
|------|-------------|
| `--clean-msocache` | Clean MSOCache files |
| `--clean-appx` | Remove AppX/MSIX packages |
| `--clean-wi-metadata` | Clean Windows Installer metadata |
| `--clean-addin-registry` | Clean add-in registry |
| `--clean-com-registry` | Clean COM registry |
| `--clean-shell-extensions` | Clean shell extensions |

### Legacy Flags

For backward compatibility, some legacy flags are still supported:

| Legacy Flag | New Syntax |
|-------------|------------|
| `--auto-all` | `remove` |
| `--repair quick` | `repair --quick` |
| `--odt-install` | `install` |
| `--diagnose` | `diagnose` |

---

## Configuration Files

Save common options in JSON:

```json
{
  "dry_run": false,
  "backup": "C:\\Backups\\Office",
  "scrub_level": "standard",
  "keep_license": true,
  "keep_templates": true,
  "passes": 2,
  "timeout": 600
}
```

Use with:
```bash
office-janitor remove --config settings.json
```

CLI flags override config file values.

---

## Safety Guidance

### Always Preview First

```bash
# Preview what will happen
office-janitor remove --dry-run --plan preview.json

# Review the plan file, then execute
office-janitor remove --backup "C:\Backups"
```

### Create Backups

```bash
# Automatic backup to specified directory
office-janitor remove --backup "C:\Backups\Office"

# System restore points are created by default
# Disable with --no-restore-point if needed
```

### Preserve User Data

```bash
# Keep everything the user might want
office-janitor remove \
  --keep-templates \
  --keep-user-settings \
  --keep-outlook-data \
  --keep-license
```

### Enterprise Deployment

```bash
# Silent unattended for SCCM/Intune
office-janitor remove --yes --quiet --no-restore-point

# Log to network share
office-janitor remove --logdir "\\server\logs\%COMPUTERNAME%"
```

---

## Logging & Diagnostics

### Log Locations

Default: `%ProgramData%\OfficeJanitor\logs`

- `human.log` – Human-readable log (rotated)
- `events.jsonl` – Machine-readable telemetry

Override with `--logdir` or `OFFICE_JANITOR_LOGDIR` environment variable.

### Diagnostics Mode

```bash
# Full diagnostic without changes
office-janitor diagnose --plan report.json -vvv

# JSON output to stdout
office-janitor diagnose --json

# Maximum verbosity
office-janitor diagnose -vvv
```

### Troubleshooting

```bash
# Skip phases to isolate issues
office-janitor remove --skip-processes --skip-services --dry-run

# Force through guardrails (use with caution)
office-janitor remove --force --skip-preflight

# Registry-only cleanup
office-janitor remove --registry-only
```

---

## TUI (Text User Interface)

Launch the interactive terminal UI:

```bash
office-janitor
```

The TUI provides:
- Live progress display with spinner
- Real-time event log
- Detection/planning/execution phases
- Key bindings for control

Auto-selects TUI when terminal supports ANSI sequences. Disable colors with `--no-color`.

---

## OffScrub Compatibility

Legacy OffScrub VBS switches are mapped to native behaviors. See `docs/CLI_COMPATIBILITY.md` for the full matrix.

```bash
# OffScrub-style flags
office-janitor offscrub --all          # Remove all
office-janitor offscrub --quiet        # Reduce output
office-janitor offscrub --test-rerun   # Double-pass
```

---

## Contributing & Testing

### Development Setup

```bash
git clone https://github.com/supermarsx/office-janitor.git
cd office-janitor
python -m venv .venv
.venv\Scripts\activate
pip install -e ".[dev]"
```

### Running Tests

```bash
# Run all tests
pytest

# With coverage
pytest --cov=src/office_janitor

# Specific test file
pytest tests/test_odt_build.py -v
```

### Code Quality

```bash
# Format
black .

# Lint
ruff check .

# Type check
mypy src tests

# All checks (PowerShell helper)
.\scripts\lint_format.ps1 -Fix
.\scripts\type_check.ps1
.\scripts\test.ps1
```

### Building

```bash
# PyInstaller executable
.\scripts\build_pyinstaller.ps1

# Distribution packages
.\scripts\build_dist.ps1
```

### Documentation Style

Use Doxygen-style docstrings:

```python
def my_function(arg: str) -> bool:
    """!
    @brief Short description.
    @details Extended description if needed.
    @param arg Description of parameter.
    @returns Description of return value.
    """
```

---

## License

See [license.md](license.md) for licensing information.

## Support

- Bug reports: Reference relevant sections of [`spec.md`](spec.md)
- Feature requests: Discuss alignment with intended architecture
- Documentation: See `docs/` for detailed guides

