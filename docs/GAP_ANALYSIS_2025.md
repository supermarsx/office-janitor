# Office Janitor - Comprehensive Gap Analysis (2025)

> **Generated:** Based on spec.md requirements, OfficeScrubber.cmd, and OffScrub*.vbs legacy scripts.
> **Last Updated:** January 2026

---

## Executive Summary

This document provides a comprehensive gap analysis comparing the current Python implementation against:
1. **spec.md** - Project specification requirements
2. **OfficeScrubber.cmd** - Batch script orchestrator (874 lines)
3. **OffScrubC2R.vbs** - C2R scrubber VBS (3,827 lines)
4. **OffScrub_O16msi.vbs** - Office 2016 MSI scrubber (4,606 lines)
5. **Other OffScrub*.vbs** - Office 2003-2013 MSI scrubbers

### Implementation Status Overview

| Category | Spec Coverage | VBS Coverage | Notes |
|----------|---------------|--------------|-------|
| Detection | ~95% | ~90% | AppX detection, vNext detection ✅ |
| MSI Uninstall | ~90% | ~85% | Setup.exe fallback, msiexec orchestration ✅ |
| C2R Uninstall | ~95% | ~90% | ODT integration, integrator.exe, license reinstall ✅ |
| License Cleanup | ~95% | ~90% | SPP/OSPP WMI, vNext registry cleanup ✅ |
| Registry Cleanup | ~95% | ~90% | 129 residue paths, TypeLib, vNext, taskband ✅ |
| File Cleanup | ~90% | ~85% | Shortcut unpinning, MSOCache, published components ✅ |
| Scheduled Tasks | ~90% | ~85% | OFFICE_SCHEDULED_TASKS_TO_DELETE constant ✅ |
| UWP/AppX | ~90% | ~85% | Detection + removal via PowerShell ✅ |
| TUI Mode | ~80% | N/A | Progress bars, hjkl navigation ✅ |
| CI/CD | ~100% | N/A | Split workflows per spec ✅ |
| Error Handling | ~90% | ~85% | Bitmask system, msiexec codes ✅ |

---

## Gap Categories

### 1. COMPLETED ITEMS ✅

#### 1.1 UWP/AppX Removal ✅ IMPLEMENTED

**Status:** COMPLETED in `fs_tools.py`
- `remove_appx_package()` - Remove single AppX package
- `remove_office_appx_packages()` - Remove all Office AppX packages
- Uses PowerShell `Remove-AppxPackage` command

---

#### 1.2 Scheduled Task Names ✅ IMPLEMENTED

**Status:** COMPLETED in `constants.py`
- `OFFICE_SCHEDULED_TASKS_TO_DELETE` - 13 task names
- `delete_office_scheduled_tasks()` in `tasks_services.py`
- Integrated into cleanup flow
"\Microsoft\Office\Office Automatic Updates"
"\Microsoft\Office\Office ClickToRun Service Monitor"
"Office Subscription Maintenance"
```

**Current State:**
- `tasks_services.py` has `delete_tasks()` function - ✅ Infrastructure exists
- **No constant with task names** - ❌ Gap

**Required Implementation:**
```python
# constants.py
OFFICE_SCHEDULED_TASKS_TO_DELETE: tuple[str, ...] = (
    "FF_INTEGRATEDstreamSchedule",
    "FF_INTEGRATEDUPDATEDETECTION",
    "C2RAppVLoggingStart",
    "Office 15 Subscription Heartbeat",
    r"\Microsoft\Office\OfficeInventoryAgentFallBack",
    r"\Microsoft\Office\OfficeTelemetryAgentFallBack",
    r"\Microsoft\Office\OfficeInventoryAgentLogOn",
    r"\Microsoft\Office\OfficeTelemetryAgentLogOn",
    "Office Background Streaming",
    r"\Microsoft\Office\Office Automatic Updates",
    r"\Microsoft\Office\Office ClickToRun Service Monitor",
    "Office Subscription Maintenance",
)
```

---

#### 1.3 CI Workflow Split Required

**Source:** spec.md Section 21 (lines 712-758)

**Spec Requires Separate Files:**
- `format.yml` - Black check
- `lint.yml` - Ruff + MyPy
- `test.yml` - Pytest matrix (Win, Python 3.9/3.11)
- `build.yml` - PyInstaller (Windows x64, optionally ARM64)
- `publish-pypi.yml` - PyPI release on tag
- `release.yml` - GitHub Release automation

**Current State:**
- Single `.github/workflows/ci.yml` (178 lines) - ❌ Not split per spec

---

### 2. HIGH PRIORITY GAPS 🟠

#### 2.1 License Menu Operations (OfficeScrubber.cmd)

**Source:** OfficeScrubber.cmd lines 724-790 (License operations submenu)

**CMD License Operations:**
| Key | Operation | Python Status |
|-----|-----------|---------------|
| C | Clean vNext Licenses | ✅ `licensing.py` |
| R | Remove ALL Licenses | ✅ `full_license_cleanup()` |
| T | Reset C2R Licenses via integrator.exe | ❌ **Missing** |
| U | Uninstall Product Keys | ✅ `licensing.py` |

**VBS Reset C2R Flow (lines 753-790):**
```batch
for %%a in (%_SKUs%) do (
    "!_Integrator!" /R /License PRIDName=%%a.16 PackageGUID="%_GUID%" PackageRoot="!_InstallRoot!"
)
```

**Required Implementation:**
```python
# c2r_uninstall.py
def reinstall_c2r_licenses(
    product_ids: Sequence[str],
    package_guid: str,
    package_root: Path,
    *,
    dry_run: bool = False,
) -> bool:
    """Re-install C2R licenses via integrator.exe /R /License."""
    ...
```

---

#### 2.2 vNext License Cache Cleanup

**Source:** OffScrubC2R.vbs `ClearVNextLicCache` (lines 1042-1050)

**VBS Implementation:**
```vbscript
Sub ClearVNextLicCache
    sLocalAppData = oWShell.ExpandEnvironmentStrings("%localappdata%")
    DeleteFolder sLocalAppData & "\Microsoft\Office\Licenses"
End Sub
```

**Current State:**
- `licensing.py` has SPP/OSPP cleanup - ✅
- `ClearVNextLicCache` folder deletion - ✅ via RESIDUE_PATH_TEMPLATES
- vNext registry cleanup - ✅ Done

**Additional vNext paths from CMD (lines 716-733):**
```batch
reg.exe delete "%kO16%\Common\Identity" /f
reg.exe delete "%kO16%\Registration" /f
reg.exe delete "%kCTR%" /f /v SharedComputerLicensing
reg.exe delete "%kCTR%" /f /v productkeys
for /f %%# in ('reg.exe query "%kCTR%" /f *.EmailAddress') do reg.exe delete "%kCTR%" /f /v %%#
for /f %%# in ('reg.exe query "%kCTR%" /f *.TenantId') do reg.exe delete "%kCTR%" /f /v %%#
for /f %%# in ('reg.exe query "%kCTR%" /f *.DeviceBasedLicensing') do reg.exe delete "%kCTR%" /f /v %%#
```

**Required Implementation:**
```python
# registry_tools.py
def cleanup_vnext_identity_registry(*, dry_run: bool = False) -> int:
    """Remove vNext identity/device licensing registry values."""
    patterns = ["*.EmailAddress", "*.TenantId", "*.DeviceBasedLicensing"]
    # Query and delete matching values
    ...
```

---

#### 2.3 Full Process Kill List

**Source:** OfficeScrubber.cmd lines 438-474

**CMD Process List (40+ processes):**
```batch
winword,excel,powerpnt,outlook,msaccess,mspub,onenote,onenotem,infopath,winproj,visio,
lync,communicator,ucmapi,groove,msosync,msouc,msoia,OneNoteM,spd,osppsvc,teams,
ose,ospd,officeC2Rclient,officeclicktorun,appvshnotify,integrator,firstrun,
OfficeHubTaskHost,msoadfsb,msoidsvcm,msoidsvc,officeondemand,integratedoffice
```

**Current State:**
- `constants.py` has `ALL_OFFICE_PROCESSES` - ✅ Most included
- Some processes may be missing from the combined list

**Verification Required:** Cross-check all 40+ processes are in the Python constants.

---

#### 2.4 User Profile Registry Loading

**Source:** OffScrubC2R.vbs `LoadUsersReg` (lines 2189-2215)

**VBS Implementation:**
```vbscript
Sub LoadUsersReg ()
    For Each profilefolder in oFso.GetFolder(sProfilesDirectory).SubFolders
        If oFso.FileExists(profilefolder.path & "\ntuser.dat") Then
            oWShell.Run "reg load " & Chr(34) & "HKU\" & profilefolder.name & Chr(34) & " " & _
                        Chr(34) & profilefolder.path & "\ntuser.dat" & Chr(34), 0, True
        End If
    Next
End Sub
```

**Current State:** ❌ **Not implemented**

**Impact:** Cannot clean per-user Office settings for all user profiles, only current user.

---

#### 2.5 Taskband Cleanup

**Source:** OffScrubC2R.vbs `ClearTaskBand` (lines 2128-2148)

**VBS Implementation:**
```vbscript
Sub ClearTaskBand ()
    sTaskBand = "Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband\"
    RegDeleteValue HKCU, sTaskBand, "Favorites", False
    RegDeleteValue HKCU, sTaskBand, "FavoritesRemovedChanges", False
    RegDeleteValue HKCU, sTaskBand, "FavoritesChanges", False
    ' ... also for all SIDs in HKU
End Sub
```

**Current State:** 
- Shortcut unpinning exists in `fs_tools.py` - ✅
- Taskband registry cleanup - ❌ **Missing**

---

### 3. MEDIUM PRIORITY GAPS 🟡

#### 3.1 TUI Mode Widget Completeness

**Source:** spec.md Section 8 (lines 170-280)

**Spec TUI Requirements:**
- [ ] Progress bars with percentage
- [ ] Scrollable pane content
- [ ] Tab switching (detection/plan/logs/settings)
- [ ] Checkbox-style plan step toggles
- [ ] Hotkey navigation (hjkl, arrows, 1-9)
- [ ] Real-time log streaming during execution

**Current State:**
- `tui.py` (1,338 lines) - Basic TUI exists
- Navigation and tabs partially implemented
- Some spec widgets may be incomplete

---

#### 3.2 OSE Service State Validation

**Source:** OffScrubC2R.vbs lines 1175-1190

**VBS Implementation:**
```vbscript
' check if OSE service is *installed, *not disabled, *running under System context.
Set OseService = oWmiLocal.Execquery("Select * From Win32_Service Where Name like 'ose%'")
For Each srvc in OseService
    If (srvc.StartMode = "Disabled") AND (Not srvc.ChangeStartMode("Manual") = 0) Then
        Log "Conflict detected: OSE service is disabled"
    If (Not srvc.StartName = "LocalSystem") AND (srvc.Change( , , , , , , "LocalSystem", "")) Then
        Log "Conflict detected: OSE service not running as LocalSystem"
Next
```

**Current State:**
- Service stop/delete exists - ✅
- Service state validation/repair before uninstall - ❌ **Missing**

---

#### 3.3 Msiexec Return Value Translation

**Source:** OffScrubC2R.vbs `SetupRetVal` function (lines 3157-3207)

**VBS Translation Table:**
```vbscript
Case 1602 : SetupRetVal = "INSTALL_USEREXIT"
Case 1603 : SetupRetVal = "INSTALL_FAILURE"
Case 1605 : SetupRetVal = "UNKNOWN_PRODUCT"
Case 1618 : SetupRetVal = "INSTALL_ALREADY_RUNNING"
Case 3010 : SetupRetVal = "SUCCESS_REBOOT_REQUIRED"
' ... 50+ error codes
```

**Current State:** ❌ **Not implemented**

**Required:** Add `MSIEXEC_RETURN_CODES` mapping to `constants.py`.

---

#### 3.4 Windows Installer REG_MULTI_SZ Cleanup

**Source:** OffScrubC2R.vbs lines 1696-1740 (Published Components cleanup)

The VBS script handles REG_MULTI_SZ values specially - it parses the multi-string, removes matching entries, and rewrites the reduced array:

```vbscript
If RegReadValue (hDefKey, sSubKeyName & item, name, sValue, "REG_MULTI_SZ") Then
    arrMultiSzValues = Split(sValue, chr(13))
    ' ... filter out matching GUIDs
    oReg.SetMultiStringValue hDefKey, sSubKeyName & item, name, arrMultiSzNewValues
End If
```

**Current State:**
- `registry_tools.py` has basic REG_MULTI_SZ support - ✅
- Selective entry removal within multi-string - ❌ **Needs verification**

---

### 4. LOW PRIORITY GAPS 🟢

#### 4.1 Logging Customizations

**Source:** OffScrubC2R.vbs `CreateLog` (lines 2690-2720)

**VBS Log Format:**
```
Microsoft Customer Support Services - Office C2R Removal Utility
Version:    2.19
64 bit OS:  True
Removal start: 14:32:15
OS Details: Windows 10 Pro, SP 0, Version: 10.0.19045...
```

**Current State:**
- `logging_ext.py` has human + JSONL logging - ✅
- Header format differs from VBS - Minor aesthetic gap

---

#### 4.2 Named Pipe Progress (LogY/LogPipe)

**Source:** OffScrubC2R.vbs `LogY`/`LogPipe` (lines 3115-3135)

**Current State:**
- `logging_ext.py` has `set_progress_pipe()`, `report_progress()`, `ProgressStages` - ✅ **Implemented**

---

#### 4.3 Return Value File Exchange

**Source:** OffScrubC2R.vbs `SetRetVal`/`GetRetValFromFile` (lines 2600-2640)

Used for elevated process return code communication.

**Current State:**
- `elevation.py` handles relaunch - ✅
- Return value file exchange - Not needed (Python handles differently)

---

## Spec.md Compliance Checklist

### Required by Spec (Section 21)

| Requirement | Status | Notes |
|-------------|--------|-------|
| format.yml (Black) | ❌ Monolithic | Needs split |
| lint.yml (Ruff + MyPy) | ❌ Monolithic | Needs split |
| test.yml (Pytest matrix) | ❌ Monolithic | Needs split |
| build.yml (PyInstaller) | ❌ Monolithic | Needs split |
| publish-pypi.yml | ❌ Missing | Not implemented |
| release.yml (GitHub Release) | ❌ Missing | Not implemented |
| requirements-dev.txt | ✅ Present | pyproject.toml dev deps |
| PyInstaller --uac-admin | ✅ Configured | office-janitor.spec |

### Detection (Spec Section 6)

| Feature | Status |
|---------|--------|
| MSI product enumeration | ✅ |
| C2R detection | ✅ |
| AppX/MSIX detection | ✅ |
| WMI-based detection | ✅ |
| Registry-based detection | ✅ |
| Version disambiguation | ✅ |

### Uninstall Orchestration (Spec Section 7)

| Feature | Status |
|---------|--------|
| C2R via ODT | ✅ |
| MSI via msiexec | ✅ |
| Setup.exe fallback | ✅ |
| Multi-pass retry | ✅ |
| Process termination | ✅ |
| Force escalation | ✅ |

### Cleanup (Spec Section 8)

| Feature | Status |
|---------|--------|
| Registry residue cleanup | ✅ |
| File/folder cleanup | ✅ |
| License cleanup (SPP/OSPP) | ✅ |
| TypeLib cleanup | ✅ |
| Shortcut cleanup | ✅ |
| WI cache orphan cleanup | ✅ |
| Scheduled task cleanup | ⚠️ Partial |
| Service cleanup | ✅ |
| **AppX/UWP removal** | ❌ Missing |

---

## Action Items Summary

### Immediate (P0)

1. **Implement UWP/AppX removal** (`uwp_uninstall.py` or `fs_tools.py`)
2. **Add scheduled task names constant** (`constants.py`)
3. **Split CI workflow files** per spec.md requirements

### Short-term (P1)

4. **Add license reset via integrator.exe** (`c2r_uninstall.py`)
5. **Add vNext identity registry cleanup** (`registry_tools.py`)
6. **Verify process kill list completeness** (`constants.py`)

### Medium-term (P2)

7. **Implement user profile registry loading** (`registry_tools.py`)
8. **Add taskband registry cleanup** (`registry_tools.py`)
9. **Add OSE service state validation** (`tasks_services.py`)
10. **Add msiexec return code translation** (`constants.py`)

### Long-term (P3)

11. **TUI widget completeness audit** (`tui.py`)
12. **REG_MULTI_SZ selective cleanup verification** (`registry_tools.py`)
13. **Add publish-pypi.yml workflow**
14. **Add release.yml workflow**

---

## Test Coverage Requirements

For each new feature, add tests to:
- `tests/test_cleanup_tools.py` - UWP removal
- `tests/test_tasks_services.py` - Scheduled tasks
- `tests/test_c2r_licensing.py` - License reset
- `tests/test_registry_tools.py` - vNext identity, taskband

---

*Document Version: 2025-01-XX*
*Based on 419 passing tests*
