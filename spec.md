# Office Janitor — Full Specification

**One‑liner:** A single, self‑contained Python 3 application (stdlib‑only, no external deps) that **detects, uninstalls, and thoroughly scrubs** Microsoft Office (MSI & Click‑to‑Run) across versions, cleans licenses/tokens, and removes residue safely — packable into a **PyInstaller one‑file** executable with UAC elevation.

---

## 1) Scope & Goals

- **Primary:** Uninstall/scrub Office 2003 → 2024/365 (MSI & C2R) on Windows 7–11, including server SKUs.
- **Secondary:** Clean orphaned activation/licensing artifacts (SPP store and OSPP), scheduled tasks, services, shortcuts, COM registrations, shell extensions, and file system leftovers.
- **Tertiary:** Provide **dry‑run**, **targeted scrub**, **auto‑all**, and **diagnostics‑only** modes with detailed logs and backups.

### Non‑Goals

- Installing or repairing Office.
- Non‑Windows platforms.

### Constraints

- **Language:** Python 3.9+ (works with 3.11+ too)
- **Dependencies:** **None** beyond Python stdlib.
- **Packaging:** PyInstaller **onefile**; include admin manifest.

---

## 2) Supported Office Variants (detection & handling)

- **MSI:** Office 2003/2007/2010/2013/2016/2019/2021 (including Visio/Project).
- **Click‑to‑Run (C2R):** Office 2013+ including Microsoft 365 Apps (Business/Enterprise), Office 2016/2019/2021/2024 Retail/Volume C2R.
- **LTS/Perp:** 2016/2019/2021/2024 perpetual.
- **Architectures:** x86/x64; mixed install detection.

---

## 3) High‑level Architecture

```
Project Root
│  office_janitor.py        # ROOT SHIM (entrypoint) — keeps imports clean
│  readme.md
│  license.md
│  pyproject.toml           # optional; not required for PyInstaller
│  .gitignore
│
├─ src/
│  └─ office_janitor/
│     __init__.py
│     main.py               # entry; arg parsing, UAC elevation check
│     detect.py             # registry & filesystem probes (MSI, C2R, app paths)
│     plan.py               # resolve requested actions → ordered plan with deps
│     scrub.py              # orchestrates uninstall, license cleanup, residue purge
│     msi_uninstall.py      # msiexec orchestration, productcode detection
│     c2r_uninstall.py      # ClickToRun orchestration (OfficeC2RClient.exe)
│     licensing.py          # SPP/OSPP token removal via PowerShell
│     registry_tools.py     # winreg helpers; export backups (.reg)
│     fs_tools.py           # path discovery, recursive delete, ACL reset
│     processes.py          # kill Office processes, stop services
│     tasks_services.py     # schtasks + sc control; uninstall leftovers
│     logging_ext.py        # structured logging (JSONL + human), rotation
│     restore_point.py      # optional system restore point creation
│     ui.py                 # CLI interactive menu (plain)
      tui.py                # **TUI mode** (ANSI/VT w/ msvcrt on Windows), spinners, widgets
│     constants.py          # GUIDs, registry roots, product mappings
│     safety.py             # dry-run, whitelist/blacklist, preflight checks
│     version.py            # __version__, build metadata
│
├─ tests/
│  ├─ __init__.py
│  ├─ test_detect.py        # registry parsing/detection unit tests (mocked)
│  ├─ test_plan.py          # action planning logic
│  ├─ test_safety.py        # guardrails, dry-run behavior
│  └─ test_registry_tools.py# export/delete simulations (temp hives/mocks)
│
└─ .github/
   └─ workflows/
      format.yml            # Black formatting check (no code changes)
      lint.yml              # Ruff + MyPy static checks
      test.yml              # Pytest on Windows only
      build.yml             # PyInstaller onefile build on Windows

# (optional) build/, dist/ — artifacts output only; not committed
```

**Root shim** (`office_janitor.py`) ensures simple launching and keeps the package under `src/`:

```python
# office_janitor.py
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))
from office_janitor.main import main

if __name__ == '__main__':
    main()
```

- **Stdlib only:** `argparse, subprocess, winreg, ctypes, json, re, shutil, pathlib, time, datetime, tempfile, zipfile, hashlib, logging, sys`.
- **Elevation:** detect admin via `ctypes.windll.shell32.IsUserAnAdmin()`. If not, **relaunch with elevation** using `ShellExecuteW("runas", ...)`.

**PyInstaller packaging** (points at the shim and provides the src path):

```
pyinstaller --onefile --uac-admin --name OfficeJanitor office_janitor.py --paths src
```

---

## 4) Operating Modes & CLI

### 4.1 CLI flags

```
Usage: office_janitor.exe [MODE] [OPTIONS]

MODE (mutually exclusive):
  --auto-all       # Detect + remove what’s present + default modern C2R set
  --target VER     # One of: 2003, 2007, 2010, 2013, 2016, 2019, 2021, 2024, 365
  --diagnose       # No changes; emit detection & plan JSON
  --cleanup-only   # No uninstalls; purge residue & licenses only

OPTIONS:
  -h, --help       # show help and exit
  -V, --version    # print version/build metadata and exit
  --include visio,project,onenote  # extend scope
  --force            # ignore certain guardrails when safe
  --dry-run          # simulate and log all steps, do not modify system
  --no-restore-point # skip creating a system restore point
  --no-license       # skip license/SPP cleanup
  --keep-templates   # keep user templates/normal.dotm, etc.
  --plan OUT.json    # write the action plan file
  --logdir DIR       # default %ProgramData%/OfficeJanitor/logs
  --backup DIR       # export relevant reg hives and files
  --timeout SEC      # global timeout for each external call
  --quiet            # minimal output (errors only)
  --json             # machine‑readable progress events to stdout
  --tui              # full‑screen TUI mode (fallbacks to plain CLI if not supported)
  --no-color         # disable ANSI colors
  --tui-compact      # TUI fits 80x24
  --tui-refresh N    # UI refresh interval ms (default 120)
```

### 4.2 Interactive (no args)

A simple text menu:

1. Detect & show installed Office
2. Auto scrub everything detected (recommended)
3. Targeted scrub (choose versions/components)
4. Cleanup only (licenses, residue)
5. Diagnostics only (export plan & inventory)
6. Settings (restore point, logging, backups)
7. Exit

### 4.3 **TUI Mode** (optional)

Enable with `--tui`. Renders a full‑screen terminal UI (stdlib‑only) using ANSI/VT sequences and `msvcrt` (Windows). If ANSI is unsupported, auto‑fallback to the plain interactive menu.

**Layout:**

- **Header bar**: app name, version, elevation status, machine.
- **Left pane**: Navigation (Detect • Plan • Uninstall • License • Cleanup • Logs • Settings).
- **Main pane (tabs):**
  - *Detect*: live inventory results (MSI/C2R/Services/Tasks) with filter box.
  - *Plan*: selectable checklist (target versions, include Visio/Project, options).
  - *Run*: step progress with spinner, per‑step status (✔/✖), elapsed time.
  - *Logs*: tail of actions JSONL + human log; PgUp/PgDn scroll.
  - *Settings*: toggles (restore point, keep templates, timeouts, logdir).

**Controls:**

- **Arrow keys / j k** navigate, **Space** toggle checkbox, **Enter** confirm, **Tab** cycle panes, **F10** start, **F1** help, **Q** quit.
- **Esc** backs out of dialogs; **/** opens inline filter in lists.

**CLI options affecting TUI:**

```
  --tui                 # force TUI
  --no-color            # disable ANSI colors
  --tui-compact         # reduced padding, fits 80x24
  --tui-refresh N       # UI refresh interval ms (default 120)
```

## **Non‑interactive automation stays identical**; TUI only changes presentation.

## 5) Detection Strategy

**MSI:**

- Query `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall` and WOW6432Node for Office products (Publisher="Microsoft Corporation" + DisplayName patterns + ProductCode GUIDs).
- Cross‑check with WMI (via `subprocess` to `wmic` or PowerShell `Get-WmiObject` fallback) when available.
- Map product codes to suites/apps (Word/Excel/Outlook/Visio/Project) via `constants.py` tables.

**C2R:**

- Check `HKLM\SOFTWARE\Microsoft\Office\ClickToRun` (Configuration, ProductReleaseIds, Updates, Client...)
- Detect `OfficeC2RClient.exe`, `integratedoffice.exe`, `root\Office16` directories.
- Enumerate installed languages, architecture, channel.

**Common:**

- Running processes (winword.exe, excel.exe, outlook.exe, onenote.exe, visio.exe, mspub.exe, teams.exe)
- Services: `ClickToRunSvc`, `osppsvc`.
- Scheduled tasks under `\Microsoft\Office\` and `\Microsoft\OfficeSoftwareProtectionPlatform\`.
- Activation state: `HKLM\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform`.

**Output:** A JSON inventory (saved under `logs/inventory-YYYYMMDD-HHMMSS.json`).

---

## 6) Uninstall Orchestration

### 6.1 MSI

- For each detected MSI productcode: `msiexec /x {GUID} /qb! /norestart` with retries.
- If DisplayIcon indicates a `setup.exe` maintenance uninstall, call directly with `/uninstall` when present.
- Wait & verify by re‑querying registry; if remnants remain, mark for residue cleanup.

### 6.2 Click‑to‑Run (C2R)

- Stop `ClickToRunSvc` and `osppsvc` safely.
- Preferred path: `OfficeC2RClient.exe /updatepromptuser=False /uninstallpromptuser=False /uninstall /displaylevel=False` when available.
- Fallback: `setup.exe` from root paths if present (`root\vfs\...`).
- Verify removal by C2R registry and file structure probes.

### 6.3 Process discipline

- Prior to uninstalls: prompt/force close Office apps. Use `taskkill /IM winword.exe /F` (progressively hardened) if user accepts.
- Suspend Outlook data providers check; warn about OST/PST integrity unaffected.

---

## 7) License / Token Cleanup

- **Goal:** mirror the effect of OffScrub license cleanup without bundling VBS.
- **SPP/OSPP:** invoke an embedded **PowerShell snippet** via `subprocess` that uses P/Invoke against `sppc.dll`/`osppc.dll` to uninstall Office license IDs (equivalent to the embedded logic observed in OffScrub helpers). The script is generated at runtime from a string literal; no external file is required.
- Remove `HKLM\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform` stale entries when safe; back up first as `.reg`.
- Restart `sppsvc`/`osppsvc` if needed.
- Optional: `cscript ospp.vbs /dstatus` probe before/after for audit (if present).

**Safety:** The license cleanup runs **only when Office binary footprints are gone** (or with `--force`).

---

## 8) Residue Cleanup

- **Filesystem:**
  - `%ProgramFiles%\Microsoft Office*`, `%ProgramFiles(x86)%\Microsoft Office*`, `%CommonProgramFiles%\Microsoft Shared\Office*`.
  - `%ProgramData%\Microsoft\Office` (cache, telemetry), `%ProgramData%\Microsoft\ClickToRun`.
  - `%LOCALAPPDATA%\Microsoft\Office`, `%APPDATA%\Microsoft\Office` (templates preserved by default).
- **Registry:**
  - `HKLM/HKCU\SOFTWARE\Microsoft\Office\*` (versioned hives), `ClickToRun`, `Common\OEM`, COM registrations under `Classes\CLSID` tied to Office.
- **Tasks/Services:** delete Office tasks (`schtasks /Delete`) and obsolete services (`sc delete` when safe).

**Guardrails:**

- Whitelist known Office paths; deny dangerous deletes outside these roots.
- Preserve user templates/Normal.dotm unless `--force`.
- Back up targeted reg keys to `backup/*.reg` before deletion.

---

## 9) Backups, Logs & Audit

- **Backups:**
  - Registry exports via `reg.exe export` for relevant hives prior to deletion.
  - Inventory and action **plan** JSON.
- **Logs (extensive, structured, trace‑rich):**
  - **Channels:**
    - **human.log** (readable text)
    - **events.jsonl** (1 line = 1 JSON event)
    - optional **stdout** JSON stream when `--json` is set.
  - **Schema (JSONL):**
    ```json
    {
      "ts":"2025-10-24T12:00:00.123Z",
      "level":"INFO|WARN|ERROR|DEBUG",
      "event":"MSIEXEC_CALL|C2R_UNINSTALL|REG_QUERY|REG_DELETE|FILE_DELETE|SERVICE_STOP|TASK_DELETE|PROCESS_KILL|LICENSE_CLEAN|PLAN_STEP|SUMMARY",
      "step_id":"uuid-or-counter",
      "args":{"cmd":"msiexec","code":"{GUID}","flags":"/x /qb! /norestart"},
      "result":{"rc":0,"duration_ms":5320},
      "corr":"session-uuid",
      "machine":{"host":"DESKTOP-123","user":"Administrator"}
    }
    ```
  - **Coverage:** **every external request/command** (msiexec, OfficeC2RClient, sc, schtasks, reg.exe, icacls, PowerShell), **every registry/file op**, and **every decision** (plan selection, guards, skip reasons) is logged.
  - **Rotation:** daily or 10 MB, keep 10 files; directory configurable via `--logdir`.
  - **Correlation IDs:** one **session UUID** + per‑step IDs; surfaced in TUI and CLI output for support.
  - **ANSI log tail in TUI:** live view of `human.log` + last N JSON events.
- **Artifacts:**
  - `inventory-*.json`, `plan-*.json`, `events-*.jsonl`, `human-*.log`, `restorepoint-*.txt`.

---

## 10) Safety & Recovery

- **Preflight:**
  - Admin check; OS version; disk free space; presence of known Office processes.
  - Offer **System Restore Point** (Windows client SKUs) via WMI/PowerShell.
- **Dry‑run:**
  - Every external call replaced by echo; file/registry ops simulated; plan emitted.
- **Rollback:**
  - Not guaranteed (uninstalls are destructive), but reg/file backups allow partial restoration.
- **Timeouts & Retries:**
  - Configurable; exponential backoff for service stops and msiexec busy states.

---

## 11) Implementation Details (stdlib only)

### 11.1 Auto‑elevation & version/help

```py
# src/office_janitor/main.py
import argparse, ctypes, sys, os
from . import version as _v

def enable_vt_mode_if_possible():
    try:
        import msvcrt  # noqa: F401
        from ctypes import wintypes
        kernel32 = ctypes.windll.kernel32
        h = kernel32.GetStdHandle(-11)  # STD_OUTPUT_HANDLE
        mode = wintypes.DWORD()
        if kernel32.GetConsoleMode(h, ctypes.byref(mode)):
            kernel32.SetConsoleMode(h, mode.value | 0x0004)
    except Exception:
        pass

def ensure_admin_and_relaunch_if_needed():
    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        is_admin = False
    if not is_admin:
        params = ' '.join(f'"{a}"' for a in sys.argv)
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
        sys.exit(0)

def build_arg_parser():
    p = argparse.ArgumentParser(prog="office-janitor", add_help=True)
    p.add_argument('-V','--version', action='version', version=f"{_v.__version__} ({_v.__build__})")
    # (modes & options added here — see CLI section)
    return p
```

### 11.2 Logger bootstrap (human + JSONL + optional stdout JSON)

```py
# src/office_janitor/logging_ext.py
import json, logging, os, sys, time, uuid, pathlib

SESSION_ID = str(uuid.uuid4())

class JsonlHandler(logging.Handler):
    def __init__(self, path):
        super().__init__()
        self.fp = open(path, 'a', encoding='utf-8')
    def emit(self, record):
        evt = {
            'ts': time.strftime('%Y-%m-%dT%H:%M:%S', time.gmtime()) + f".{int(record.msecs):03d}Z",
            'level': record.levelname,
            'event': getattr(record, 'event', record.msg.split(':',1)[0][:32]),
            'step_id': getattr(record, 'step_id', None),
            'corr': SESSION_ID,
            'msg': record.getMessage(),
        }
        if hasattr(record, 'data'):
            evt['data'] = record.data
        self.fp.write(json.dumps(evt, ensure_ascii=False) + '
')
        self.fp.flush()

def setup_logging(root_dir, json_to_stdout=False, level=logging.INFO):
    path = pathlib.Path(root_dir)
    path.mkdir(parents=True, exist_ok=True)
    human = logging.getLogger('human')
    machine = logging.getLogger('machine')
    human.setLevel(level)
    machine.setLevel(level)
    # human log
    hf = logging.FileHandler(path / 'human.log', encoding='utf-8')
    hf.setFormatter(logging.Formatter('[%(asctime)s] %(levelname)s: %(message)s'))
    human.addHandler(hf)
    # JSONL
    machine.addHandler(JsonlHandler(path / 'events.jsonl'))
    # optional JSON to stdout
    if json_to_stdout:
        machine.addHandler(logging.StreamHandler(sys.stdout))
    return human, machine
```

### 11.3 TUI engine (no third‑party deps) (no third‑party deps)

**Module:** `tui.py`

- **Rendering:** ANSI escape sequences + box‑drawing chars; double‑buffers by composing lines then flushing.
- **Input:** `msvcrt.getwch()` on Windows; (POSIX fallback via `termios`/`tty` only if ever needed).
- **Loop:** \~8 FPS (configurable); reacts to an internal event queue.
- **Widgets:** `Label`, `Menu`, `CheckboxList`, `LogPane` (file tail), `Dialog` (confirmations), `Spinner`.
- **Threading:** long‑running actions emit progress via a thread‑safe `Queue` consumed by the renderer.
- **Graceful fallback:** if VT mode can’t be enabled or stdout isn’t a TTY, auto‑fallback to plain CLI.

**Sketch:**

```py
# src/office_janitor/tui.py
import sys, os, time, queue
try:
    import msvcrt
except ImportError:
    msvcrt = None

CSI = "["
RESET = CSI + "0m"
BOLD = CSI + "1m"
REV = CSI + "7m"

class TUI:
    def __init__(self, app):
        self.app = app
        self.events = queue.Queue()
        self.running = True
        self.compact = False
        self.refresh_ms = 120

    def run(self):
        self.clear()
        while self.running:
            self.draw()
            self.poll_keys()
            time.sleep(self.refresh_ms/1000)

    def clear(self):
        sys.stdout.write(CSI + "2J" + CSI + "H")
        sys.stdout.flush()

    def draw(self):
        # compose header + panes; use app state (inventory/plan/progress)
        sys.stdout.write(CSI + "H")
        sys.stdout.write(BOLD + " Office Janitor "+ RESET + "  [F1 Help] [F10 Run] [Q Quit]
")
        # ... render nav + main pane ...
        sys.stdout.flush()

    def poll_keys(self):
        if msvcrt and msvcrt.kbhit():
            ch = msvcrt.getwch()
            # map to actions (arrows, enter, space, etc.)
            if ch in ('q', 'Q'):
                self.running = False
```

`main.py` decides between plain CLI and TUI:

```py
# src/office_janitor/main.py
import argparse, sys
from .tui import TUI

def parse_args():
    p = argparse.ArgumentParser()
    modes = p.add_mutually_exclusive_group()
    modes.add_argument('--auto-all', action='store_true')
    modes.add_argument('--target')
    modes.add_argument('--diagnose', action='store_true')
    modes.add_argument('--cleanup-only', action='store_true')
    p.add_argument('--tui', action='store_true')
    p.add_argument('--no-color', action='store_true')
    p.add_argument('--tui-compact', action='store_true')
    p.add_argument('--tui-refresh', type=int, default=120)
    # ... other flags ...
    return p.parse_args()

def main():
    ensure_admin_and_relaunch_if_needed()
    enable_vt_mode_if_possible()
    args = parse_args()
    if args.tui and sys.stdout.isatty():
        TUI(app_state).run()  # app_state wired to detection/plan/scrub engine
    else:
        run_plain_cli(args)
```

---

## 12) Packaging (PyInstaller)

- **Spec:** onefile, console, uac‑admin manifest (redundant with runtime elevation but helps).
- **Command:**
  ```bash
  pyinstaller --onefile --uac-admin --name OfficeJanitor office_janitor.py --paths src
  ```
- **Artifacts:**
  - **Windows x64** primary.
  - **Windows ARM64** optional via a **self‑hosted** Windows on ARM runner (PyInstaller builds natively on target arch).
- **Version info:** embed from `version.py`.
- **Signing (optional):** document `signtool` usage for Windows.

---

## 13) UX Copy & Prompts

- Clear confirmation before destructive actions:
  - “This will remove Microsoft Office and related licensing artifacts from this machine. Continue? (Y/n)”
- Post‑run summary: counts of removed products, files, reg keys; where logs live.
- Suggest reboot if msiexec or C2R returns reboot‑required.

---

## 14) Testing Matrix

- **OS:** Win7 SP1, Win10 22H2, Win11 24H2, Server 2012R2/2016/2019/2022.
- **Office:**
  - MSI: 2007, 2010, 2013, 2016, 2019, 2021 (Visio/Project variants)
  - C2R: 2016/2019/2021/2024/365 (x86/x64, different channels)
- **Scenarios:** mixed installs, partial components, language packs, pending updates, running apps, low‑privilege start → elevation, offline machines.

---

## 15) Failure Modes & Handling

- **msiexec busy / error 1618:** queue and retry; prompt to close Windows Installer sessions.
- **ClickToRunSvc stop timeout:** escalate: disable service → reboot advise.
- **Files locked:** schedule delete on reboot via `MoveFileEx` (ctypes) or `PendingFileRenameOperations` registry fallback.
- **PowerShell blocked by policy:** use `-ExecutionPolicy Bypass`; if refused, warn & skip license cleanup unless `--force`.

---

## 16) Mapping vs OffScrub family (reference)

- Equivalent coverage for: Office 2003/2007/2010/2013/2016/2019/2021/2024 & 365; MSI & C2R; OSPP/SPP license removal.
- No redistribution of VBS; logic implemented natively or via short embedded PS as text.
- Optional compatibility flags mirroring OffScrub switches (e.g., `--target 2013`, `--include visio`).

---

## 17) Telemetry & Privacy

- **Default:** telemetry OFF. No network calls. All operations local.
- **Optional:** local machine‑readable logs for enterprise ingestion.

---

## 18) Future Extensions

- GUI wrapper (tkinter) with the same engine.
- Signed driverless file unlock (Restart Manager API via ctypes for fewer reboots).
- Offline cache parser for Office C2R builds to free disk.

---

## 19) Deliverables

- Source tree under `src/` with root shim `office_janitor.py`.
- `.github/workflows/*.yml` for format, lint, test, build.
- `OfficeJanitor.exe` (one‑file), `SHA256.txt`.
- `readme.md` with usage, `license.md`.

---

## 22) Reference Code Integration (office-janitor-draft-code / OfficeScrubber)

Use the provided draft assets as **behavioral references** to shape our Python components. We **do not** redistribute VBS/CMD content; we re‑implement equivalent logic using stdlib + PowerShell bridges.

### 22.1 Source artifacts observed

- `OfficeScrubber/OfficeScrubberAIO.cmd` — master orchestrator for scrub flows.
- `OfficeScrubber/OfficeScrubber.cmd` — entry‑menu wrapper.
- `OfficeScrubber/bin/OffScrub03.vbs` — Office 2003 MSI scrub helper.
- `OfficeScrubber/bin/OffScrub07.vbs` — Office 2007 MSI scrub helper.
- `OfficeScrubber/bin/OffScrub10.vbs` — Office 2010 MSI scrub helper.
- `OfficeScrubber/bin/OffScrub_O15msi.vbs` — Office 2013 MSI scrub helper.
- `OfficeScrubber/bin/OffScrub_O16msi.vbs` — Office 2016+ MSI scrub helper.
- `OfficeScrubber/bin/OffScrubC2R.vbs` — Click‑to‑Run scrub helper.
- `OfficeScrubber/bin/CleanOffice.txt` — PowerShell license removal (defines `UninstallLicenses($dll)` using P/Invoke)
- `OfficeScrubber/README.md` — overview of versions/flows.

### 22.2 Mapping → Python modules

| Reference asset                    | Purpose                                       | Python counterpart                                                     |
| ---------------------------------- | --------------------------------------------- | ---------------------------------------------------------------------- |
| OfficeScrubberAIO.cmd              | End‑to‑end orchestration & menu               | `scrub.py` (orchestrator), `ui.py` / `tui.py` (menus)                  |
| OffScrub03/07/10/O15msi/O16msi.vbs | MSI variant uninstalls & residue rules        | `msi_uninstall.py` + `detect.py` + `fs_tools.py` + `registry_tools.py` |
| OffScrubC2R.vbs                    | C2R uninstall & cleanup                       | `c2r_uninstall.py`                                                     |
| CleanOffice.txt (PS)               | SPP/OSPP license token uninstall via P/Invoke | `licensing.py` (embeds, generates, and runs PS)                        |
| OfficeScrubber.cmd                 | menu launcher                                 | `main.py` + `ui.py` / `tui.py`                                         |

### 22.3 Functional parity checklist

- **Detection parity**: replicate registry/product code probing patterns used by OffScrub family; ensure 32/64‑bit hives and WOW6432 coverage.
- **Uninstall ordering**: follow AIO sequencing (apps → suites → languages → shared components) to minimize repair/reinstall churn.
- **C2R removal**: prefer `OfficeC2RClient.exe /uninstall` flags as in reference; stop `ClickToRunSvc` first.
- **License cleanup**: port `UninstallLicenses($DllPath)` (from CleanOffice.txt) semantics using embedded PowerShell with dynamic P/Invoke.
- **Residue rules**: mirror directory & registry deletion lists implied by OffScrub scripts; apply our guardrails/whitelist.
- **User data**: conform to reference’s stance—preserve user templates and data by default.

### 22.4 Port notes (license removal)

`CleanOffice.txt` defines a PowerShell function that dynamically P/Invokes `SLOpen`, `SLGetSLIDList`, and `SLUninstallLicense` from **SPP DLLs**, targeting Office SLIDs (e.g., `0ff1ce15-…`). Our `licensing.py` generates a self‑contained PS script at runtime to:

1. Open SPP handle (`SLOpen`).
2. Enumerate Office SLIDs via `SLGetSLIDList` with the Office app GUID filter.
3. Loop and call `SLUninstallLicense` per SLID.
4. Repeat against `osppc.dll` where applicable.
5. Emit before/after counts to logs.

### 22.5 Python stubs influenced by reference

```py
# src/office_janitor/msi_uninstall.py
from .logging_ext import setup_logging
from .registry_tools import find_msi_products
from .fs_tools import guarded_delete

MSI_SILENT_FLAGS = ['/x', '/qb!', '/norestart']

def uninstall_msi_product(product_code, timeout=1800):
    # Mirrors OffScrub *msi* behaviors: silent UI, no restart
    cmd = ['msiexec.exe'] + MSI_SILENT_FLAGS.copy()
    cmd[1] = cmd[1] + f' {{{product_code}}}' if cmd[1] == '/x' else cmd[1]
    return run_command(cmd, timeout=timeout, event='MSIEXEC_CALL', data={'product': product_code})
```

```py
# src/office_janitor/c2r_uninstall.py
C2R_CLIENTS = [
    r"C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe",
    r"C:\Program Files (x86)\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe",
]

C2R_ARGS = ['/updatepromptuser=False','/uninstallpromptuser=False','/uninstall','/displaylevel=False']

def uninstall_c2r():
    # Stop services like reference scripts do
    stop_service('ClickToRunSvc')
    for exe in C2R_CLIENTS:
        if path_exists(exe):
            return run_command([exe] + C2R_ARGS, timeout=3600, event='C2R_UNINSTALL')
    return {'rc': 2, 'error': 'OfficeC2RClient.exe not found'}
```

```py
# src/office_janitor/licensing.py
PS_TEMPLATE = r'''
function UninstallLicenses($DllPath) {
  $TB = [AppDomain]::CurrentDomain.DefineDynamicAssembly(4,1).DefineDynamicModule(2).DefineType(0)
  [void]$TB.DefinePInvokeMethod('SLOpen', $DllPath, 22, 1, [int], @([IntPtr].MakeByRefType()), 1, 3)
  [void]$TB.DefinePInvokeMethod('SLGetSLIDList', $DllPath, 22, 1, [int], @([IntPtr],[int],[Guid].MakeByRefType(),[int],[int].MakeByRefType(),[IntPtr].MakeByRefType()), 1, 3).SetImplementationFlags(128)
  [void]$TB.DefinePInvokeMethod('SLUninstallLicense', $DllPath, 22, 1, [int], @([IntPtr],[IntPtr]), 1, 3)
  $SPPC = $TB.CreateType(); $Handle = 0; [void]$SPPC::SLOpen([ref]$Handle)
  $pnReturnIds = 0; $ppReturnIds = 0
  if (!$SPPC::SLGetSLIDList($Handle, 0, [ref][Guid]'0ff1ce15-0000-0000-0000-000000000000', 6, [ref]$pnReturnIds, [ref]$ppReturnIds)) {
    foreach ($i in 0..($pnReturnIds - 1)) { [void]$SPPC::SLUninstallLicense($Handle, [Int64]$ppReturnIds + ($i*16)) }
  }
}
UninstallLicenses('sppc.dll')
'''
```

> **Note:** GUID above is placeholder; `constants.py` will include the exact Office SLID filters mirrored from the reference scripts.

### 22.6 Tests derived from reference behavior

- Ensure MSI uninstall calls are formed like OffScrub: `/x {GUID} /qb! /norestart`.
- Ensure C2R args match OffScrubC2R: no prompts, no UI.
- Ensure licensing PS emits uninstall count > 0 when mock SLIDs are present.

### 22.7 Risk considerations

- VBS logic sometimes contains special‑cases for SKUs/locales; we will codify these as data in `constants.py` and guard behind `--force` where destructive.
- Some flows require multiple passes; `scrub.py` will iterate until detection re‑probes clean or a pass limit is reached.

---

## 20) Acceptance Criteria

- On a machine with Office C2R 2019 and Visio Pro MSI 2016, `--auto-all` removes both, cleans licenses, leaves user templates intact, and produces inventory/plan/logs. Post‑reboot, no Office apps run or appear in Add/Remove, and license status is clean.

---

## 21) CI & Quality Gates

**Philosophy:** runtime has zero external deps; CI may use dev tools.

### 21.1 `format.yml` (Black check)

- **Trigger:** push/PR.
- **Job:** ubuntu‑latest → `pip install black` → `black --check src tests office_janitor.py`.

### 21.2 `lint.yml` (Ruff + MyPy)

- **Trigger:** push/PR.
- **Jobs:**
  - `ruff`: `pip install ruff` → `ruff check src tests`.
  - `mypy`: `pip install mypy types-ctypes` → `mypy --ignore-missing-imports src`.

### 21.3 `test.yml` (Pytest)

- **Matrix:** `os: [windows-latest]`, `python: [3.9, 3.11]`.
- **Steps:** checkout → setup‑python → `pip install -r requirements-dev.txt` → `pytest -q`.

### 21.4 `build.yml` (PyInstaller, Windows only)

- **Trigger:** push to `main` and tags `v*`.
- **Matrix:**
  - **Windows:** `windows-latest` (x64).
  - *(Optional)* add a **self‑hosted** Windows ARM64 runner for an ARM64 artifact.
- **Key steps:**
  - Setup Python 3.11.
  - `pip install pyinstaller`.
  - Build: `pyinstaller --onefile --uac-admin --name OfficeJanitor office_janitor.py --paths src`.
  - Upload artifacts (e.g., `OfficeJanitor-win-x64.exe`, optionally `OfficeJanitor-win-arm64.exe`).

`build.yml` (PyInstaller, multi‑OS/arch)

- **Trigger:** push to `main` and tags `v*`.
- **Matrix:**
  - **Windows:** `windows-latest` (x64), `windows-2022-arm64` (or self‑hosted arm64).
  - **macOS:** `macos-13` (x86\_64), `macos-14` (arm64).
  - **Linux:** `ubuntu-latest` (x86\_64), aarch64 via `uraimo/run-on-arch-action@v2`.
- **Key steps:**
  - Setup Python 3.11.
  - `pip install pyinstaller`.
  - Build: `pyinstaller --onefile --uac-admin --name OfficeJanitor office_janitor.py --paths src`.
  - Upload artifacts per‑platform (e.g., `OfficeJanitor-win-x64.exe`, `OfficeJanitor-macos-arm64`, etc.).

### 21.5 `publish-pypi.yml` (PyPI when everything passes)

- **Trigger:** on **tag **`` and **workflow\_run** success for `format`, `lint`, `test`, `build`.
- **Build dist:** `python -m build` → creates `sdist` + wheel (Windows‑targeted package; functionality is Windows‑only).
- **Twine upload:** `python -m twine upload dist/*` using `PYPI_API_TOKEN` secret.
- **Package metadata:** console entry point `office-janitor=office_janitor.main:main` via `pyproject.toml`; classifiers indicate **Operating System :: Microsoft :: Windows**.

### 21.6 `release.yml` (Auto GitHub Release)

- **Trigger:** `workflow_run` after `format`, `lint`, `test`, `build` **all succeed** for a tag.

- **Actions:** create GitHub Release named from tag; attach Windows artifacts + checksums. `release.yml` (Auto GitHub Release)

- **Trigger:** `workflow_run` after `format`, `lint`, `test`, `build` **all succeed** for a tag.

- **Actions:** create GitHub Release named from tag; attach all PyInstaller artifacts + checksums.

### 21.7 Suggested `requirements-dev.txt`

```
black
ruff
mypy
pytest
pyinstaller
build
twine
```

> Note: These are **dev‑only** and do not affect the app’s stdlib‑only runtime.

