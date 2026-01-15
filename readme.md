# Office Janitor

<img width="1017" height="902" alt="image" src="https://github.com/user-attachments/assets/37748fbb-f3a6-446b-81ec-1c2780e7137b" />

<br>
<br>

Office Janitor is a stdlib-only Python utility that inventories, uninstalls, and scrubs Microsoft Office across MSI and Click-to-Run deployments. The tool follows the architecture defined in [`spec.md`](spec.md) and can be packaged into a single-file Windows executable with PyInstaller.

## Installation

### Prerequisites
- Windows 7 or later with administrator privileges for uninstall operations.
- Python 3.9 or later.
- Optional developer tooling: `pip`, `virtualenv`, and the CI-aligned utilities (`black`, `ruff`, `mypy`, `pytest`, `pyinstaller`).

### Running from source
1. Clone the repository and enter the project directory.
2. Create and activate a virtual environment (optional but recommended).
3. Install the package in editable mode:
   ```bash
   python -m pip install --upgrade pip
   python -m pip install -e .
   ```
4. Launch the shim entry point:
   ```bash
   python office_janitor.py --help
   ```

### Building the PyInstaller executable
Office Janitor is designed to ship as a one-file executable. After installing PyInstaller, build the binary with:
```bash
pyinstaller --onefile --uac-admin --name office-janitor office_janitor.py --paths src
```
The resulting executable will appear in `dist/office-janitor.exe`.

## Usage

### Quick start
- Detect and scrub everything automatically:
  ```bash
  python office_janitor.py --auto-all
  ```
- Run a diagnostic dry run and export logs/plan without making changes:
  ```bash
  python office_janitor.py --diagnose --dry-run --plan diagnostics/plan.json
  ```
- Launch the interactive menu when no mode flag is supplied:
  ```bash
  python office_janitor.py
  ```

### OffScrub compatibility
Legacy OffScrub VBS switches are mapped to native behaviours; see `docs/CLI_COMPATIBILITY.md` for the full matrix (e.g., `/PREVIEW`, `/QUIET`, `/NOREBOOT`, reruns, license/cache skips). Return codes preserve the reboot bit and honour `/RETERRORORSUCCESS` when requested.

### Running with the module path
If the shim is unavailable, the same interface can be invoked via the package module:
```bash
python -m office_janitor.main --auto-all
```

## CLI Reference

### Modes
| Flag | Description |
| --- | --- |
| `--auto-all` | Detect, plan, and scrub everything that matches the supported Office footprints. |
| `--target VER` | Focus on a specific Office release (`2003`, `2007`, `2010`, `2013`, `2016`, `2019`, `2021`, `2024`, `365`). |
| `--diagnose` | Perform detection and planning only; write inventory/plan artifacts. |
| `--cleanup-only` | Skip uninstallers; remove residue, scheduled tasks, services, and licensing debris. |

Exactly one mode can be chosen at a time. When no mode is provided the CLI falls back to the interactive menu defined in `ui.py`.

### Core options
- `-h, --help` – display contextual help.
- `-V, --version` – emit version and build metadata from `version.py`.
- `--include visio,project,onenote` – extend the cleanup scope.
- `--dry-run` – simulate all actions without mutating the system.
- `--force` – bypass non-critical guardrails described in `safety.py`.
- `--allow-unsupported-windows` – permit execution on Windows releases below the supported minimum.
- `--plan OUT.json` – persist the generated plan to disk.
- `--logdir DIR` – override the log directory (defaults described below).
- `--backup DIR` – export registry/file backups to the specified directory.
- `--timeout SEC` – set per-step timeout in seconds.
- `--no-license` – skip SPP/OSPP license cleanup.
- `--keep-license` – preserve Office licenses (alias of `--no-license`).
- `--keep-templates` – preserve user templates like `normal.dotm`.
- `--no-restore-point` – prevent automatic system restore point creation.
- `--limited-user` – run detection and uninstall stages under a limited user token when possible.
- `--json` – mirror structured progress events to stdout in addition to `events.jsonl`.
- `--tui`, `--tui-compact`, `--tui-refresh N` – control the terminal UI.
- `--no-color` – disable ANSI/VT colours when terminals misbehave.
- `--quiet` – reduce human-readable logging noise to errors only.

## TUI Overview
The optional Textual User Interface (TUI) mirrors the CLI capabilities using ANSI/VT sequences. Launch it explicitly with `--tui` or allow Office Janitor to auto-select it when a compatible terminal is detected. The TUI displays:
- A header banner with version/build metadata and elevation status.
- Live progress panes for detection, planning, safety checks, and execution steps.
- Event history sourced from the same queue consumed by the CLI (`ui_events`).
- Key bindings for pausing updates or exiting gracefully.

If the environment cannot render ANSI sequences, the application automatically falls back to the plain interactive menu.

## Logging Locations
Logging is configured by `logging_ext.setup_logging` and produces both human-readable and JSONL output:
- Human log: `human.log` (rotated, plain text).
- Machine log: `events.jsonl` (structured telemetry).

By default logs are written beneath `%ProgramData%/OfficeJanitor/logs` on Windows. The directory can be overridden via `--logdir` or the `OFFICE_JANITOR_LOGDIR` environment variable. When no explicit location exists (e.g., non-Windows development systems) the tool falls back to `./logs` or the XDG state directory. Each run emits metadata including a session UUID, timestamps, and build identifiers.

## Safety Guidance
- **Dry runs first:** Always begin with `--dry-run` or `--diagnose` to review the planned actions and confirm guardrails from `safety.perform_preflight_checks` pass.
- **Review plans:** Persist plans with `--plan` and inspect them before executing destructive steps, especially in enterprise environments.
- **Target carefully:** Use `--target` or `--cleanup-only` to avoid removing unintended Office versions or user data. Templates and customisations can be preserved with flags like `--keep-templates`.
- **Back up the registry:** Provide `--backup` (or rely on automatic restore points) so the registry state can be restored if needed.
- **Run elevated:** Uninstallation and cleanup require administrative rights; the CLI auto-prompts for elevation on Windows but you should verify the consent dialogue.
- **Monitor logs:** Consult `human.log` and `events.jsonl` for detailed progress, warnings, and error diagnostics after each run.

## Contributing & Testing
1. Fork and clone the repository, then install it in editable mode as shown above.
2. Adhere to the Doxygen-style docstring convention (`"""!` with `@brief`/`@details`) for modules, classes, and public callables.
3. Before opening a pull request, run the same checks enforced by CI:
   ```bash
   # Format and lint
   black .
   ruff check .
   mypy src tests

   # Execute tests
   pytest

   # Validate packaging
   pyinstaller --onefile --uac-admin --name OfficeJanitor office_janitor.py --paths src
   ```
4. Ensure new features include or update tests beneath `tests/` and keep the documentation in sync with the behaviour changes.
5. Confirm logs and backups generated by new code continue to respect the defaults in `fs_tools.get_default_log_directory` and related helpers.

### Dev helper scripts (PowerShell, Windows-friendly)
- `scripts/lint_format.ps1 [-Fix]` — run Ruff + Black (use `-Fix` to apply changes).
- `scripts/type_check.ps1` — run mypy with the repo configuration.
- `scripts/test.ps1 [-- pytest args]` — run pytest, forwarding any extra arguments.
- `scripts/build_dist.ps1 [-RefreshTools]` — build sdist/wheel artifacts; pass `-RefreshTools` to ensure `build` is installed.
- `scripts/build_pyinstaller.ps1` — package the admin-elevated single-file executable.

Bug reports and feature requests should reference the relevant sections of [`spec.md`](spec.md) so the discussion remains aligned with the intended architecture.

## ODT XML Configuration Builder

Office Janitor includes a powerful Office Deployment Tool (ODT) XML configuration generator and installer. This allows you to create installation, removal, and download configurations for any Office product, and execute installations directly without needing separate tools.

### ODT Installation (New!)

You can now install Office directly using the embedded ODT setup.exe:

```bash
# Install Office LTSC 2024 + Visio + Project with 3 languages
office-janitor --odt-install --odt-preset ltsc2024-full-x64 \
  --odt-language en-us --odt-language es-mx --odt-language pt-br

# Install clean version (no OneDrive/Skype bloat)
office-janitor --odt-install --odt-preset ltsc2024-full-x64-clean \
  --odt-language en-us --odt-language de-de

# Preview what would be installed (dry-run)
office-janitor --odt-install --odt-preset office2024-x64 --dry-run

# Install Microsoft 365 Apps without OneDrive and Skype
office-janitor --odt-install --odt-preset 365-proplus-x64-clean
```

### Listing Available Options

```bash
# List all available Office products
office-janitor --odt-list-products

# List pre-built configuration presets
office-janitor --odt-list-presets

# List update channels
office-janitor --odt-list-channels

# List supported languages
office-janitor --odt-list-languages
```

### Quick Install Aliases

For convenience, author-defined shortcuts are available for common installations:

```bash
# Goobler: LTSC 2024 + Visio Pro + Project Pro (no OneDrive/Skype) with pt-pt and en-us
office-janitor --goobler

# Pupa: LTSC 2024 ProPlus only (no OneDrive/Skype) with pt-pt and en-us
office-janitor --pupa

# Both support --dry-run to preview
office-janitor --goobler --dry-run
```

| Alias | Preset | Products | Languages |
|-------|--------|----------|-----------|
| `--goobler` | `ltsc2024-full-x64-clean` | ProPlus 2024 + Visio Pro + Project Pro | pt-pt, en-us |
| `--pupa` | `ltsc2024-x64-clean` | ProPlus 2024 only | pt-pt, en-us |

### Building Installation Configurations

```bash
# Use a preset (simplest method)
office-janitor --odt-build --odt-preset 365-proplus-x64 --odt-output install.xml

# Office LTSC 2024 with German and French
office-janitor --odt-build --odt-preset office2024-x64 \
  --odt-language de-de --odt-language fr-fr --odt-output install.xml

# Custom product selection
office-janitor --odt-build --odt-product O365ProPlusRetail \
  --odt-include-visio --odt-include-project --odt-output install.xml

# 32-bit Office 2021 for shared computers
office-janitor --odt-build --odt-preset office2021-x86 \
  --odt-shared-computer --odt-output install.xml

# Exclude OneDrive and Teams from M365
office-janitor --odt-build --odt-product O365ProPlusRetail \
  --odt-exclude-app OneDrive --odt-exclude-app Teams --odt-output install.xml
```

### Building Removal Configurations

```bash
# Remove all Office products
office-janitor --odt-removal --odt-output remove.xml

# Remove specific products only
office-janitor --odt-removal --odt-product VisioProRetail \
  --odt-product ProjectProRetail --odt-output remove.xml

# Remove all and clean MSI leftovers
office-janitor --odt-removal --odt-remove-msi --odt-output remove.xml
```

### Building Download Configurations

```bash
# Download Office for offline installation
office-janitor --odt-download "C:\ODTOffline" \
  --odt-preset 365-proplus-x64 --odt-output download.xml

# Download multiple languages for deployment
office-janitor --odt-download "D:\OfficeSource" \
  --odt-product O365ProPlusRetail \
  --odt-language en-us --odt-language es-es --odt-language pt-br \
  --odt-output download.xml
```

### Complete Preset Reference

#### Microsoft 365 (Subscription-based)

| Preset | Products | Description |
|--------|----------|-------------|
| `365-proplus-x64` | O365ProPlusRetail | Microsoft 365 Apps for enterprise (64-bit) |
| `365-proplus-x86` | O365ProPlusRetail | Microsoft 365 Apps for enterprise (32-bit) |
| `365-business-x64` | O365BusinessRetail | Microsoft 365 Apps for business (64-bit) |
| `365-proplus-visio-project` | O365ProPlusRetail, VisioProRetail, ProjectProRetail | M365 Apps + Visio + Project (64-bit) |
| `365-shared-computer` | O365ProPlusRetail | M365 Apps with Shared Computer Licensing |
| `365-proplus-x64-clean` | O365ProPlusRetail | M365 Apps - **No OneDrive/Skype** |
| `365-proplus-visio-project-clean` | O365ProPlusRetail, VisioProRetail, ProjectProRetail | M365 + Visio + Project - **No OneDrive/Skype** |

#### Office LTSC 2024 (Perpetual/Volume License)

| Preset | Products | Description |
|--------|----------|-------------|
| `office2024-x64` | ProPlus2024Volume | Office LTSC 2024 Professional Plus (64-bit) |
| `office2024-x86` | ProPlus2024Volume | Office LTSC 2024 Professional Plus (32-bit) |
| `office2024-standard-x64` | Standard2024Volume | Office LTSC 2024 Standard (64-bit) |
| `ltsc2024-full-x64` | ProPlus2024Volume, VisioPro2024Volume, ProjectPro2024Volume | **Office 2024 + Visio Pro + Project Pro (64-bit)** |
| `ltsc2024-full-x86` | ProPlus2024Volume, VisioPro2024Volume, ProjectPro2024Volume | Office 2024 + Visio Pro + Project Pro (32-bit) |
| `ltsc2024-x64-clean` | ProPlus2024Volume | Office 2024 ProPlus - **No OneDrive/Skype** |
| `ltsc2024-full-x64-clean` | ProPlus2024Volume, VisioPro2024Volume, ProjectPro2024Volume | **Office 2024 Full Suite - No OneDrive/Skype** |
| `ltsc2024-full-x86-clean` | ProPlus2024Volume, VisioPro2024Volume, ProjectPro2024Volume | Office 2024 Full Suite (32-bit) - No bloat |

#### Office LTSC 2021 (Perpetual/Volume License)

| Preset | Products | Description |
|--------|----------|-------------|
| `office2021-x64` | ProPlus2021Volume | Office LTSC 2021 Professional Plus (64-bit) |
| `office2021-x86` | ProPlus2021Volume | Office LTSC 2021 Professional Plus (32-bit) |
| `office2021-standard-x64` | Standard2021Volume | Office LTSC 2021 Standard (64-bit) |
| `ltsc2021-full-x64` | ProPlus2021Volume, VisioPro2021Volume, ProjectPro2021Volume | Office 2021 + Visio Pro + Project Pro (64-bit) |

#### Office 2019 (Perpetual/Volume License)

| Preset | Products | Description |
|--------|----------|-------------|
| `office2019-x64` | ProPlus2019Volume | Office 2019 Professional Plus (64-bit) |
| `office2019-x86` | ProPlus2019Volume | Office 2019 Professional Plus (32-bit) |

#### Standalone Products

| Preset | Products | Description |
|--------|----------|-------------|
| `visio-pro-x64` | VisioPro2024Volume | Visio Professional 2024 (64-bit) |
| `project-pro-x64` | ProjectPro2024Volume | Project Professional 2024 (64-bit) |

### Supported Languages

Office Janitor supports 60+ language codes. Common languages include:

| Code | Language | Code | Language |
|------|----------|------|----------|
| `en-us` | English (US) | `de-de` | German |
| `en-gb` | English (UK) | `fr-fr` | French (France) |
| `es-es` | Spanish (Spain) | `fr-ca` | French (Canada) |
| `es-mx` | Spanish (Mexico) | `it-it` | Italian |
| `pt-br` | Portuguese (Brazil) | `ja-jp` | Japanese |
| `pt-pt` | Portuguese (Portugal) | `ko-kr` | Korean |
| `zh-cn` | Chinese (Simplified) | `ru-ru` | Russian |
| `zh-tw` | Chinese (Traditional) | `pl-pl` | Polish |
| `ar-sa` | Arabic | `nl-nl` | Dutch |
| `he-il` | Hebrew | `tr-tr` | Turkish |

Use `office-janitor --odt-list-languages` for the complete list.

### Update Channels

| Channel | Value | Use Case |
|---------|-------|----------|
| Current | `Current` | Latest features (M365 subscription) |
| Monthly Enterprise | `MonthlyEnterprise` | Monthly updates with support (M365) |
| Semi-Annual | `SemiAnnual` | Stability-focused, twice yearly (M365) |
| Perpetual VL 2024 | `PerpetualVL2024` | Office LTSC 2024 |
| Perpetual VL 2021 | `PerpetualVL2021` | Office LTSC 2021 |
| Perpetual VL 2019 | `PerpetualVL2019` | Office 2019 |
| Beta | `BetaChannel` | Preview/Insider builds |

## Common Command Combinations

### ODT Installation Workflows

```bash
# Install Office LTSC 2024 + Visio + Project with multiple languages
office-janitor --odt-install --odt-preset ltsc2024-full-x64 \
  --odt-language en-us --odt-language es-mx --odt-language pt-br

# Install clean Office 2024 (no OneDrive, no Skype/Lync)
office-janitor --odt-install --odt-preset ltsc2024-full-x64-clean \
  --odt-language en-us

# Install Microsoft 365 for enterprise without bloatware
office-janitor --odt-install --odt-preset 365-proplus-x64-clean \
  --odt-language en-us --odt-language de-de

# Install just ProPlus 2024 (no Visio/Project)
office-janitor --odt-install --odt-preset office2024-x64 --odt-language en-us

# Custom installation with specific products
office-janitor --odt-install --odt-product ProPlus2024Volume \
  --odt-product VisioPro2024Volume --odt-channel PerpetualVL2024 \
  --odt-language en-us --odt-exclude-app OneDrive --odt-exclude-app Lync

# Preview installation without executing (dry-run)
office-janitor --odt-install --odt-preset ltsc2024-full-x64 --dry-run
```

### Diagnostics & Planning

```bash
# Full diagnostic without making changes
office-janitor --diagnose --dry-run --plan report.json --verbose

# Export detection results and plan to JSON
office-janitor --diagnose --plan plan.json --json

# Verify what would be removed before actual run
office-janitor --auto-all --dry-run --plan preview.json
```

### Safe Removal Workflows

```bash
# Step 1: Always preview first
office-janitor --auto-all --dry-run --plan step1_plan.json

# Step 2: Review the plan, then execute with backup
office-janitor --auto-all --backup "C:\Backups\OfficeRemoval" --yes

# Remove everything but keep user data and licenses
office-janitor --auto-all --keep-templates --keep-user-settings --keep-license

# Aggressive cleanup after failed uninstalls
office-janitor --cleanup-only --scrub-level aggressive --clean-all-licenses
```

### Targeted Version Removal

```bash
# Remove only Office 2016
office-janitor --target 2016 --backup "C:\Backups"

# Remove Office 365/Microsoft 365 only
office-janitor --target 365 --dry-run

# Remove Office 2019 including Visio and Project
office-janitor --target 2019 --include visio,project
```

### Click-to-Run Operations

```bash
# Quick repair of Office C2R
office-janitor --repair quick

# Full online repair
office-janitor --repair full --repair-visible

# Use custom repair configuration
office-janitor --repair-config "C:\Configs\custom_repair.xml"

# Remove C2R with ODT
office-janitor --use-odt --auto-all
```

### MSI Office Operations

```bash
# Remove MSI Office only
office-janitor --msi-only --auto-all

# Remove specific MSI product by GUID
office-janitor --product-code "{90160000-0011-0000-0000-0000000FF1CE}"

# Multiple passes for stubborn MSI installs
office-janitor --target 2010 --passes 3
```

### Enterprise Deployment

```bash
# Silent unattended removal (for scripts/SCCM/Intune)
office-janitor --auto-all --yes --quiet --no-restore-point

# Generate removal config for ODT deployment
office-janitor --odt-removal --odt-remove-msi --odt-output remove_all.xml

# Cleanup only mode for post-uninstall scripts
office-janitor --cleanup-only --clean-all-licenses --clean-shortcuts

# Export logs to network share
office-janitor --auto-all --logdir "\\server\logs\%COMPUTERNAME%"
```

### License Management

```bash
# Clean all Office licenses
office-janitor --cleanup-only --clean-all-licenses

# Clean SPP tokens only (KMS/MAK)
office-janitor --cleanup-only --clean-spp

# Clean vNext/device-based licensing
office-janitor --cleanup-only --clean-vnext

# Preserve licenses during removal
office-janitor --auto-all --keep-license
```

### Troubleshooting

```bash
# Maximum verbosity for debugging
office-janitor --diagnose -vvv --json

# Skip specific phases to isolate issues
office-janitor --auto-all --skip-processes --skip-services --dry-run

# Registry-only cleanup (skip uninstall/filesystem)
office-janitor --registry-only

# Force through guardrails (use with caution)
office-janitor --auto-all --force --skip-preflight
```

### Using Configuration Files

```bash
# Use a saved configuration
office-janitor --config enterprise_removal.json

# Override config options from CLI
office-janitor --config base.json --dry-run --verbose

# Combine config with specific target
office-janitor --config cleanup_settings.json --target 2019
```

