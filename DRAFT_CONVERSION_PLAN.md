# Draft Code Conversion Plan

## Goal
Convert all code in `office-janitor-draft-code/` into Python modules under `src/office_janitor/`, integrate or replace existing scripts, add tests, and delete the draft folder after verification.

## Inventory (completed)
- `office-janitor-draft-code/OfficeScrubber.cmd` (main batch menu/orchestration)
- `office-janitor-draft-code/OfficeScrubberAIO.cmd` (AIO variant of batch orchestration)
- `office-janitor-draft-code/bin/OffScrub_O16msi.vbs` (VBScript helper)
- `office-janitor-draft-code/bin/OffScrub_O15msi.vbs`
- `office-janitor-draft-code/bin/OffScrubC2R.vbs`
- `office-janitor-draft-code/bin/OffScrub10.vbs`
- `office-janitor-draft-code/bin/OffScrub07.vbs`
- `office-janitor-draft-code/bin/OffScrub03.vbs`
- `office-janitor-draft-code/bin/CleanOffice.txt` (embedded PowerShell payload)
- `office-janitor-draft-code/README.md`

## Mapping (source → target)
- Batch orchestrator (`OfficeScrubber.cmd`, `OfficeScrubberAIO.cmd`) → integrate into `src/office_janitor/scrub.py` and `src/office_janitor/tui.py` where interactive menu and orchestration live.
- OffScrub VBScript helpers (`bin/*.vbs`) → represented by `src/office_janitor/off_scrub_scripts.py` which already contains shim generation; ensure the embedded shim bodies match originals where needed or provide compatibility wrappers.
- `CleanOffice.txt` embedded PowerShell → port the logic into `src/office_janitor/licensing.py` as a safe, unit-testable function to uninstall licenses via `ctypes` or controlled PowerShell invocation. For non-Windows CI, keep a mocked path and unit tests that assert call parameters.
- README → merge relevant user-facing docs into project `README.md` or `docs/` if present.

## Prioritization
1. High: Orchestration logic in `OfficeScrubber.cmd` → `scrub.py` integration (affects end-to-end behavior). Estimate: 2-3 days.
2. High: `OffScrubC2R.vbs` and C2R uninstall flows → ensure `c2r_uninstall.py` uses `off_scrub_scripts.build_offscrub_command` as appropriate. Estimate: 1 day.
3. Medium: `CleanOffice.txt` license removal → integrate into `licensing.py` with safe flags. Estimate: 1-2 days.
4. Low: Old MSI OffScrub scripts placeholders — ensure `off_scrub_scripts._SCRIPT_BODIES` contain adequate shims. Estimate: 0.5 day.

## Detailed Implementation Steps
1. Update `off_scrub_scripts._SCRIPT_BODIES` where necessary to mirror original scripts' command-line signatures and comments. Verify `ensure_offscrub_script` behaviour.
2. Port `OfficeScrubber.cmd` menu logic into Python:
   - Reuse existing `detect.py` detection functions instead of registry queries in batch file.
   - Expose a CLI entrypoint in `office_janitor.py` or `src/office_janitor/main.py` to accept equivalent flags (`/A`, `/C`, `/M6`, etc.) mapped to argparse flags (`--all`, `--c2r`, `--m16`, ...).
   - Implement an interactive TUI path in `tui.py` reusing existing functions; replicate menu choices as Python functions.
   - Maintain safety checks: require `--elevate` or check `is_admin()` before destructive ops.
3. Port `CleanOffice.txt` payload:
   - Implement a function `licensing.uninstall_licenses_with_dll(dll_path: Path)` that uses `ctypes` to call SLOpen/SLGetSLIDList/SLUninstallLicense on Windows, and falls back to PowerShell invocation for safety.
   - Add robust error handling and require `--force` to run.
4. Integrate OffScrub flows:
   - Ensure `scrub.py` calls `build_offscrub_command(kind='msi', version=...)` and runs via `exec_utils.run_command`.
   - Remove direct calls to `cscript`/batch files from Python; use generated shim scripts only for compatibility when external tools expect them.
5. Tests:
   - Add unit tests for the new CLI flags and mapping, mocking system interactions.
   - Add tests for `licensing.uninstall_licenses_with_dll` that mock `ctypes`/PowerShell.
6. CI:
   - Run full test suite. Mark Windows-only integration tests to run in a Windows environment.
7. Cleanup:
   - After all tests pass and manual Windows smoke tests done, delete `office-janitor-draft-code/` with `git rm -r` in a separate commit.

## Safety & Validation
- All destructive functions must honour `--dry-run`.
- Require explicit confirmation or `--force` for license uninstall.
- Keep original draft folder until after merge and CI green.

## Commands to run (local)
- Run tests: `python -m pytest -q`
- Run CLI smoke test: `python office_janitor.py --dry-run --all`
- Install editable: `python -m pip install -e .`

## Timeline Estimate
- Total: ~1–2 weeks, depending on Windows testing availability.

## Next Action
- Start by ensuring `off_scrub_scripts._SCRIPT_BODIES` reflect the draft VBS signatures. I will update shim bodies where needed and add a mapping doc to the repo.


