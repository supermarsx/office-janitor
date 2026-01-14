# TODO

> **See also:** 
> - [docs/GAP_ANALYSIS_2025.md](docs/GAP_ANALYSIS_2025.md) - Comprehensive 2025 gap analysis
> - [docs/SCRUBBER_GAP_ANALYSIS.md](docs/SCRUBBER_GAP_ANALYSIS.md) - Original VBS-to-Python feature parity analysis

---

## NEW GAPS (2025 Analysis)

### Critical Gaps 🔴

- [x] **UWP/AppX Removal** (`fs_tools.py`) ✅ COMPLETED
  - [x] `remove_appx_package(package_full_name)` - Remove single AppX package
  - [x] `remove_office_appx_packages()` - Remove all Office AppX packages
  - [x] Use PowerShell `Remove-AppxPackage` (CMD line 476)
  - [x] Add tests for dry-run and live removal (5 tests added)
  - **Note:** Detection already exists in `detect.py:detect_appx_packages()`

- [x] **Scheduled Task Names Constant** (`constants.py`) ✅ COMPLETED
  - [x] Add `OFFICE_SCHEDULED_TASKS_TO_DELETE` tuple with 13 task names
  - [x] Wire into `tasks_services.delete_tasks()` during cleanup
  - [x] Added `delete_office_scheduled_tasks()` convenience function

### High Priority Gaps 🟠

- [x] **C2R License Reset via Integrator** (`c2r_uninstall.py`) ✅ COMPLETED
  - [x] `get_c2r_product_release_ids()` - Query active SKUs from registry
  - [x] `get_c2r_install_root()` - Get install root and package GUID
  - [x] `reinstall_c2r_license()` - Reinstall single SKU license
  - [x] `reinstall_c2r_licenses()` - Reinstall all detected SKU licenses
  - [x] Uses `integrator.exe /R /License PRIDName=... PackageGUID=...`
  - [x] Based on OfficeScrubber.cmd license menu option T (lines 753-790)

- [x] **vNext Identity Registry Cleanup** (`registry_tools.py`) ✅ COMPLETED
  - [x] `cleanup_vnext_identity_registry(dry_run=False)`
  - [x] Delete values matching: `*.EmailAddress`, `*.TenantId`, `*.DeviceBasedLicensing`
  - [x] Delete keys: `Common\Identity`, `Registration`, `SharedComputerLicensing`
  - [x] Based on OfficeScrubber.cmd lines 716-733

- [x] **Process Kill List Verification** (`constants.py`) ✅ COMPLETED
  - [x] Cross-checked all 40+ processes from OfficeScrubber.cmd (lines 438-474)
  - [x] Added 12 missing processes to `C2R_INFRASTRUCTURE_PROCESSES`
  - [x] Added: werfault, mstore, setlang, ois, graph, OfficeHubTaskHost, msoidsvc, msoidsvcm, ucmapi, sdxhelper, OfficeClickToRun, officec2rclient

- [x] **User Profile Registry Loading** (`registry_tools.py`) ✅ COMPLETED
  - [x] `get_user_profiles_directory()` - Get %USERPROFILE% parent path
  - [x] `get_user_profile_hive_paths()` - Find ntuser.dat paths
  - [x] `load_user_registry_hives()` - Load all ntuser.dat files to HKU
  - [x] `unload_user_registry_hives()` - Cleanup after processing
  - [x] `get_loaded_user_hives()` - Track loaded hives
  - [x] Based on OffScrubC2R.vbs `LoadUsersReg` (lines 2189-2215)

- [x] **Taskband Registry Cleanup** (`registry_tools.py`) ✅ COMPLETED
  - [x] `cleanup_taskband_registry(include_all_users, dry_run=False)`
  - [x] Delete Favorites, FavoritesRemovedChanges, FavoritesChanges, FavoritesResolve, FavoritesVersion
  - [x] Process for HKCU and all SIDs in HKU (when include_all_users=True)
  - [x] Based on OffScrubC2R.vbs `ClearTaskBand` (lines 2128-2148)

### Medium Priority Gaps 🟡

- [x] **OSE Service State Validation** (`tasks_services.py`) ✅ COMPLETED
  - [x] `validate_ose_service_state(dry_run=False)` - Check OSE service before uninstall
  - [x] Enable if disabled, set to LocalSystem if wrong account
  - [x] Based on OffScrubC2R.vbs lines 1175-1190

- [x] **Msiexec Return Code Translation** (`constants.py`) ✅ COMPLETED
  - [x] Add `MSIEXEC_RETURN_CODES` dict mapping 50+ codes to names
  - [x] Add `translate_msiexec_return_code(code)` function
  - [x] Based on OffScrubC2R.vbs `SetupRetVal` (lines 3157-3207)

- [x] **TUI Widget Audit** (`tui.py`) ✅ COMPLETED
  - [x] Progress bars with percentage display (`render_progress_bar()`)
  - [x] Scrollable pane content with hjkl navigation (vim-style keys)
  - [x] Real-time log streaming during execution
  - [x] Checkbox-style plan step toggles

### Low Priority Gaps 🟢

- [x] **REG_MULTI_SZ Selective Cleanup** (`registry_tools.py`) ✅ COMPLETED
  - [x] `filter_multi_string_value()` - Filter entries from REG_MULTI_SZ
  - [x] `cleanup_published_components()` - Clean Office entries from WI Published Components
  - [x] `is_office_guid()` - Detect Office GUIDs matching VBS InScope() logic
  - [x] `_decode_squished_guid()` - Decode Windows Installer compressed GUIDs
  - [x] Based on OffScrubC2R.vbs lines 1696-1740

- [x] **CI Workflows Split** (`.github/workflows/`) ✅ COMPLETED
  - [x] `format.yml` - Black formatting check
  - [x] `lint.yml` - Ruff + MyPy static checks
  - [x] `test.yml` - Pytest on Windows matrix
  - [x] `build.yml` - PyInstaller onefile + distributions
  - [x] `publish-pypi.yml` - PyPI publishing
  - [x] `release.yml` - GitHub Release automation

---

## Legacy VBS Scrubber Implementation Gaps (COMPLETED)

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

- [x] **OSPP License Cleanup** (`licensing.py`) ✅ COMPLETED
  - [x] WMI query for `SoftwareLicensingProduct` (Win8+)
  - [x] WMI query for `OfficeSoftwareProtectionProduct` (Win7)
  - [x] `UninstallProductKey()` method call
  - [x] Filter by Office ApplicationId: `0ff1ce15-a989-479d-af46-f275c6370663`
  - [x] vNext cache cleanup, activation tokens, SCL cache
  - [x] `full_license_cleanup()` orchestration function

- [x] **TypeLib Cleanup** (`registry_tools.py`) ✅ COMPLETED
  - [x] Add ~17 Office TypeLib GUIDs to constants
  - [x] Scan `HKLM\\Software\\Classes\\TypeLib\\{GUID}`
  - [x] Check if target file exists
  - [x] Remove orphaned registrations

- [x] **ODT Integration** (`c2r_uninstall.py`) ✅ COMPLETED
  - [x] `build_remove_xml()` - generate RemoveAll.xml config
  - [x] `download_odt(version)` - fetch from Microsoft CDN
  - [x] `find_or_download_odt()` - locate local or download
  - [x] `uninstall_via_odt()` - execute setup.exe /configure

- [x] **Full File Cleanup** (expand `constants.py` RESIDUE_PATH_TEMPLATES) ✅ COMPLETED
  - [x] C2R package folders (Microsoft Office 15/16, root, PackageManifests)
  - [x] Common Files\\Microsoft Shared\\ClickToRun
  - [x] Common Files\\Microsoft Shared\\OFFICE15/16
  - [x] User data folders (AppData, LocalAppData Office subfolders)
  - [x] ProgramData\\Microsoft\\ClickToRun

- [x] **MSOCache Cleanup** (`fs_tools.py`) ✅ COMPLETED
  - [x] Scan all fixed drive roots for MSOCache
  - [x] Filter by product code patterns
  - [x] Delete only folders for products being removed

### Phase 3: Medium Priority 🟡

- [x] **Setup.exe Based Uninstall** (`msi_uninstall.py`) ✅ COMPLETED
  - [x] Locate setup.exe from InstallSource/InstallLocation registry
  - [x] Build uninstall config XML
  - [x] Execute maintenance mode removal before msiexec fallback

- [x] **Shortcut Unpinning** (`fs_tools.py`) ✅ COMPLETED
  - [x] Use PowerShell Shell.Application for verb discovery
  - [x] Find "Unpin from taskbar" / "Unpin from Start" verbs
  - [x] Execute unpinning before shortcut deletion

- [x] **Integrator.exe Invocation** (`c2r_uninstall.py`) ✅ COMPLETED
  - [x] Delete C2RManifest*.xml files
  - [x] Call integrator.exe /U with PackageRoot and PackageGUID
  - [x] `find_c2r_package_guids()` - discover packages from registry
  - [x] `unregister_all_c2r_integrations()` - batch unregistration

- [x] **WI Cache Orphan Cleanup** (`fs_tools.py`) ✅ COMPLETED
  - [x] Scan %WINDIR%\\Installer for .msi/.msp files
  - [x] Check if product code is still registered
  - [x] Delete orphaned installer cache files

- [x] **Service Management** (`tasks_services.py` + `constants.py`) ✅ COMPLETED
  - [x] `delete_services()` - net stop + sc delete (already existed)
  - [x] `OFFICE_SERVICES_TO_DELETE` constant added
  - [x] Target: OfficeSvc, ClickToRunSvc, ose, ose64, osppsvc

### Phase 4: Low Priority 🟢

- [x] **Named Pipe Progress Reporting** (`logging_ext.py`) ✅ COMPLETED
  - [x] `set_progress_pipe()` / `get_progress_pipe()` - configuration
  - [x] `report_progress(stage)` - VBS LogY equivalent
  - [x] `ProgressStages` class - standard stage identifiers

- [x] **Error Bitmask System** (`constants.py`) ✅ COMPLETED
  - [x] `ScrubErrorCode` IntFlag enum matching VBS constants
  - [x] SUCCESS, FAIL, REBOOT_REQUIRED, USER_CANCEL, etc.

- [x] **Explorer Restart** (`processes.py`) ✅ COMPLETED
  - [x] `is_explorer_running()` - check if running
  - [x] `restart_explorer_if_needed()` - check and restart if terminated

- [ ] **Temp ARP Entry Creation** (`detect.py`) ✅ COMPLETED
  - [x] For orphaned products without configuration entry
  - [x] Create temporary HKLM\\...\\Uninstall\\OFFICE_TEMP.xxx keys
  - [x] `find_orphaned_wi_products()` - scan WI metadata
  - [x] `create_temp_arp_entry()` - create single entry
  - [x] `cleanup_temp_arp_entries()` - remove temp entries
  - [x] `create_arp_entries_for_orphans()` - batch creation

### Phase 5: Product Classification

- [x] **MSI Product Type Classification** (`guid_utils.py`) ✅ COMPLETED
  - [x] Add `OFFICE_PRODUCT_TYPE_CODES` mapping (000F, 0011, 0012, etc.)
  - [x] `get_product_type_code(product_code)` - extract type code
  - [x] `classify_office_product(product_code)` - return suite/single/server/c2r
  - [x] Support for all Office product GUID patterns

---

## Test Coverage for New Features

- [x] `tests/test_msi_components.py` - Component scanning with mocked WI COM ✅
- [x] `tests/test_guid_utils.py` - GUID compression/expansion algorithms ✅
- [x] `tests/test_registry_tools.py` - WI metadata validation and cleanup ✅
- [x] `tests/test_fs_tools.py` - File system tools including MSOCache ✅
- [x] `tests/test_cleanup_tools.py` - Cleanup orchestration ✅
- [x] `tests/test_c2r_licensing.py` - WMI-based license cleanup ✅
- [x] `tests/test_uninstallers.py` - Integrator.exe/ODT tests ✅
- [x] `tests/test_detect.py` - Temp ARP entry management ✅

---

## Notes

- VBS scripts total ~26,450 lines across 6 files
- Python implementation currently covers ~60-70% of detection, ~50% of cleanup
- Most critical gap: Windows Installer component scanning (required for thorough MSI cleanup)
- Keep `--dry-run` support for all new functionality per spec.md
z