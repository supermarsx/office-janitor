# OffScrub VBS -> Python TODO

## Scope & Source Material
- VBS originals live under `office-janitor-draft-code/bin/*.vbs` and `out_vbs_chunks/*.txt`. Treat them as behavioural references while keeping new code in `src/office_janitor/` per `spec.md`.
- Target Python modules:
  - CLI/compat surface: `src/office_janitor/off_scrub_native.py`, `src/office_janitor/off_scrub_scripts.py`
  - Uninstall orchestration: `scrub.py`, `c2r_uninstall.py`, `msi_uninstall.py`
  - Detection/planning hooks: `detect.py`, `plan.py`, `tui.py`
  - Safety/logging/licensing: `safety.py`, `logging_ext.py`, `licensing.py`

## Cross-Cutting Tasks
- [ ] Build a complete CLI compatibility table for all VBS switches (C2R + MSI). Map each to Python semantics (e.g., `/QUIET` -> quiet logging, `/NOREBOOT` -> suppress reboot prompts, `/PREVIEW` -> dry-run). *Docs still needed; behaviour largely mapped.*
- [x] Normalise legacy invocation parsing so `python -m office_janitor.off_scrub_native` accepts VBS-style `/` flags plus the script path arguments that legacy `cscript` uses.
- [x] Tie legacy flags to data-driven behaviours (keep-license/offline/skip-shortcut-detection/no-reboot/quiet/reruns) and forward them to installers/cleanup.
- [x] Detect-target selection: when `ALL` or a version-specific script is invoked, automatically select matching MSI/C2R installs from `detect.gather_office_inventory()` rather than requiring explicit product codes.
- [ ] Logging parity: emit human + JSONL records that mirror OffScrub stages (stage0 detection, stage1 uninstall, rerun markers) and include legacy return-code semantics.
- [x] Preserve guardrails: ensure all destructive actions honour `--dry-run`/`/PREVIEW` and `safety` checks; require explicit confirmation for license removal paths.
- [x] Add limited-user mode flag and de-elevation path so detection/uninstalls can run under non-admin tokens when requested.

## Script-Specific Porting Tasks
- [ ] **OffScrubC2R.vbs**
  - Capture switches `/DETECTONLY|/PREVIEW`, `/OFFLINE|/FORCEOFFLINE`, `/NOREBOOT`, `/QUIET|/PASSIVE`, `/RETERRORORSUCCESS`, `/KL|/KEEPLICENSE`, `/LOG`, `/SKIPSD`, `/FORCEARPUNINSTALL`, `/TESTRERUN`.
  - Map uninstall flows to `c2r_uninstall.uninstall_products`, adding forced ARP uninstall fallback and rerun/backoff semantics; ensure service/task cleanup aligns with `tasks_services` helpers.
  - [x] Port Click-to-Run scheduled task cleanup into data-driven cleanup steps; COM compatibility/registry residue cleanup added; ClickToRun cache directories now removed. (COM cache purge still pending.)
- [ ] **OffScrub_O16msi.vbs / OffScrub_O15msi.vbs**
  - Handle MSI options `/OSE`, `/ENDCURRENTINSTALLS`, `/DELETEUSERSETTINGS`, `/CLEARADDINREG`, `/REMOVELYNC`, `/KEEPUSERSETTINGS`, `/FASTREMOVE`, `/BYPASS`, `/SCANCOMPONENTS`, `/REMOVEOSPP`, `/NOREBOOT`, `/QUIET|/PASSIVE`.
  - Align product selection with `constants.MSI_UNINSTALL_VERSION_GROUPS` so 2016/2019/2021/2024 resolve to the same stage; reuse detection metadata to populate uninstall handles and setup.exe fallbacks.
  - Port user-profile cleanup (basic Start Menu shortcut purges added; full detection/unpinning still pending) plus add-in registry and VBA cleanup into `fs_tools`/`registry_tools` with dry-run support.
- [ ] **OffScrub10.vbs / OffScrub07.vbs / OffScrub03.vbs**
  - Mirror legacy flags listed above (minus OSPP handling) and ensure detection filters target the correct major versions (11->2003, 12->2007, 14->2010).
  - Port legacy task/service cleanup (OSE, Groove/OneDrive, VC runtime prompts) into structured cleanup steps shared with newer scripts.
- [ ] **Shared returns**
  - Recreate legacy return code bitmask (success, reboot required, user cancel, stage failures) and expose a translation layer so Python exit codes remain batch-compatible.

## Testing & Validation
- [x] Add unit tests covering legacy argument parsing, version-group selection from script names, and inventory-based target selection (mocked detection).
- [x] Add integration-style tests that stub uninstall functions to assert correct call parameters for representative flag combinations.
- [x] Add tests for limited-user flag propagation and VBA/add-in cleanup hooks.
- [ ] Document parity gaps and any intentionally unimplemented switches in `README.md`/module docstrings.

## References
- Draft conversion notes: `DRAFT_CONVERSION_PLAN.md`
- Original batch/TUI orchestration expectations: `office-janitor-draft-code/OfficeScrubber*.cmd`, `spec.md`
