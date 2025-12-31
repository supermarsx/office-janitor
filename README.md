# Office Janitor

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
pyinstaller --onefile --uac-admin --name OfficeJanitor office_janitor.py --paths src
```
The resulting executable will appear in `dist/OfficeJanitor.exe`.

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
- `--plan OUT.json` – persist the generated plan to disk.
- `--logdir DIR` – override the log directory (defaults described below).
- `--backup DIR` – export registry/file backups to the specified directory.
- `--no-license` – skip SPP/OSPP license cleanup.
- `--no-restore-point` – prevent automatic system restore point creation.
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
