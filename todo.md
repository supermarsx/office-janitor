# TODO

> **See also:** [docs/SCRUBBER_GAP_ANALYSIS.md](docs/SCRUBBER_GAP_ANALYSIS.md) for detailed VBS-to-Python feature parity analysis.

---

## Legacy VBS Scrubber Implementation Gaps

### Phase 1: Critical (Must Have) 🔴

- [x] **MSI Component Scanner** (`msi_components.py` - NEW FILE) ✅ COMPLETED
  - [x] Create `WindowsInstaller` COM wrapper class
  - [x] Implement `enumerate_products()` - list all MSI product codes
  - [x] Implement `enumerate_components()` - list all WI component GUIDs  
  - [x] Implement `get_component_clients(component_id)` - products owning component
  - [x] Implement `get_component_path(product, component)` - file/registry keypath
  - [x] Implement `MSIComponentScanner` class matching VBS `ScanComponents`
  - [x] Generate FileList.txt, RegList.txt, CompVerbose.txt logs

- [x] **GUID Utilities** (`guid_utils.py` - NEW FILE) ✅ COMPLETED
  - [x] `compress_guid()` - convert `{GUID}` to 32-char compressed form
  - [x] `expand_guid()` - reverse of compress
  - [x] `squish_guid()` - 20-char squished format
  - [x] `decode_squished_guid()` - reverse of squish
  - [x] Office product type classification

- [x] **WI Metadata Validation** (`registry_tools.py`) ✅ COMPLETED
  - [x] `validate_wi_metadata_key(hive, key, expected_length)` - validate entries
  - [x] `scan_wi_metadata()` - scan all standard WI paths
  - [x] `cleanup_wi_orphaned_products()` - remove orphaned product entries
  - [x] `cleanup_wi_orphaned_components()` - remove orphaned component entries

- [x] **Full Registry Cleanup** (expanded `constants.py`) ✅ COMPLETED
  - [x] Add all C2R registry paths from OffScrubC2R.vbs `RegWipe`
  - [x] Add all MSI registry paths from OffScrub_O16msi.vbs `RegWipe`
  - [x] Add shell integration paths (protocol handlers, overlays, etc.)
  - [x] Add add-in registration paths
  - [x] Add service registry paths
  - [x] Expanded from ~50 to 129 registry residue paths

- [x] **Shell Integration Cleanup** (`registry_tools.py`) ✅ COMPLETED
  - [x] `cleanup_orphaned_typelibs()` - remove orphaned TypeLib registrations
  - [x] `cleanup_protocol_handlers()` - remove orphaned protocol handlers
  - [x] `cleanup_shell_extensions()` - scan shell extension approvals
  - [x] TypeLib GUIDs added to constants (17 Office TypeLibs)

### Phase 2: High Priority 🟠

- [ ] **OSPP License Cleanup** (`licensing.py`)
  - [ ] WMI query for `SoftwareLicensingProduct` (Win8+)
  - [ ] WMI query for `OfficeSoftwareProtectionProduct` (Win7)
  - [ ] `UninstallProductKey()` method call
  - [ ] Filter by Office ApplicationId: `0ff1ce15-a989-479d-af46-f275c6370663`

- [ ] **TypeLib Cleanup** (`registry_tools.py`)
  - [ ] Add ~80 Office TypeLib GUIDs to constants
  - [ ] Scan `HKLM\Software\Classes\TypeLib\{GUID}`
  - [ ] Check if target file exists
  - [ ] Remove orphaned registrations

- [ ] **ODT Integration** (`c2r_uninstall.py`)
  - [ ] `build_remove_xml()` - generate RemoveAll.xml config
  - [ ] `download_odt(version)` - fetch from Microsoft CDN
  - [ ] `uninstall_via_odt()` - execute setup.exe /configure

- [ ] **Full File Cleanup** (expand `constants.py` RESIDUE_PATH_TEMPLATES)
  - [ ] C2R package folders (Microsoft Office 15/16, root, PackageManifests)
  - [ ] Common Files\Microsoft Shared\ClickToRun
  - [ ] Common Files\Microsoft Shared\OFFICE15/16
  - [ ] User data folders (AppData, LocalAppData Office subfolders)
  - [ ] ProgramData\Microsoft\ClickToRun

- [ ] **MSOCache Cleanup** (`fs_tools.py`)
  - [ ] Scan all fixed drive roots for MSOCache
  - [ ] Filter by product code patterns
  - [ ] Delete only folders for products being removed

### Phase 3: Medium Priority 🟡

- [ ] **Setup.exe Based Uninstall** (`msi_uninstall.py`)
  - [ ] Locate setup.exe from InstallSource/InstallLocation registry
  - [ ] Build uninstall config XML
  - [ ] Execute maintenance mode removal before msiexec fallback

- [ ] **Shortcut Unpinning** (`fs_tools.py`)
  - [ ] Use Shell.Application COM for verb discovery
  - [ ] Find "Unpin from taskbar" / "Unpin from Start" verbs
  - [ ] Execute unpinning before shortcut deletion

- [ ] **Integrator.exe Invocation** (`c2r_uninstall.py`)
  - [ ] Delete C2RManifest*.xml files
  - [ ] Call integrator.exe /U with PackageRoot and PackageGUID

- [ ] **WI Cache Orphan Cleanup** (`fs_tools.py`)
  - [ ] Scan %WINDIR%\Installer for .msi/.msp files
  - [ ] Check if product code is still registered
  - [ ] Delete orphaned installer cache files

- [ ] **Service Management** (`tasks_services.py`)
  - [ ] `delete_service(name)` - net stop + sc delete
  - [ ] Target: OfficeSvc, ClickToRunSvc, ose, ospp

### Phase 4: Low Priority 🟢

- [ ] **Named Pipe Progress Reporting** (`logging_ext.py`)
  - [ ] `NamedPipeReporter` class for external monitoring
  - [ ] Stage markers: stage0, stage1, CleanOSPP, reboot, ok

- [ ] **Error Bitmask System** (`constants.py`)
  - [ ] `ScrubErrorCode` IntFlag enum matching VBS constants
  - [ ] Apply throughout scrub orchestration

- [ ] **Explorer Restart** (`processes.py`)
  - [ ] `restart_explorer_if_needed()` - check and restart if terminated

- [ ] **Temp ARP Entry Creation** (`detect.py`)
  - [ ] For orphaned products without configuration entry
  - [ ] Create temporary HKLM\...\Uninstall\OFFICE15.xxx keys

### Phase 5: Product Classification

- [ ] **MSI Product Type Classification** (`constants.py` + `detect.py`)
  - [ ] Add `MSI_PRODUCT_TYPE_CODES` mapping (000F, 0011, 0012, etc.)
  - [ ] `classify_msi_product(product_code)` - return suite/single/server/c2r
  - [ ] Support for Lync/Skype product GUID list

---

## Existing Technical Debt

- [ ] Mypy compliance (currently ~151 errors):
  - [ ] `plan.py`: pass mutable sequences into `_augment_auto_all_c2r_inventory`, remove unused ignores, replace `int/list/dict` casts on `object` with validated conversions.
  - [ ] UI/TUI: add typed app_state/ logger references, type event queues (`deque[dict[str, object]]`), clean unused ignores, and ensure `msvcrt` stub typing is acceptable.
  - [ ] Scrub/Plan orchestration: fix `build_plan` input types (`dict[str, Sequence[dict]]`), scrub result typing, avoid `dict(obj)`/`list(obj)` on unknowns, ensure `c2r_uninstall`/`msi_uninstall` interfaces are typed.
  - [ ] Safety/registry/fs tools: remove unused `type: ignore` comments; ensure winreg stubs cover attr-defined errors; add explicit `Mapping` types where `.get` is used.
  - [ ] Detect/licensing/off_scrub: validate collections before casting to `list`/`dict`; resolve duplicate variable names and assignment types.
  - [ ] Logging: `_SizedTimedRotatingFileHandler` args typed; remove unused ignores.

- [ ] CI workflows: split monolithic `.github/workflows/ci.yml` into `format.yml`, `lint.yml`, `test.yml`, `build.yml`, `publish-pypi.yml`, and `release.yml` per spec.

---

## Test Coverage for New Features

- [ ] `tests/test_msi_components.py` - Component scanning with mocked WI COM
- [ ] `tests/test_guid_utils.py` - GUID compression/expansion algorithms
- [ ] `tests/test_wi_cleanup.py` - WI metadata validation and cleanup
- [ ] `tests/test_shell_cleanup.py` - Shell integration registry cleanup
- [ ] `tests/test_typelib_cleanup.py` - Orphaned TypeLib detection
- [ ] `tests/test_ospp_cleanup.py` - WMI-based license removal
- [ ] `tests/test_mso_cache.py` - MSOCache cleanup logic

---

## Notes

- VBS scripts total ~26,450 lines across 6 files
- Python implementation currently covers ~60-70% of detection, ~50% of cleanup
- Most critical gap: Windows Installer component scanning (required for thorough MSI cleanup)
- Keep `--dry-run` support for all new functionality per spec.md
