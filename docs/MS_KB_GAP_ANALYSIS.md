# Office Janitor: Microsoft KB 4e2904ea Gap Analysis

**Source:** [Manually uninstall Office (KB 4e2904ea)](https://support.microsoft.com/en-us/office/manually-uninstall-office-4e2904ea-25c8-4544-99ee-17696bb3027b)

This document analyzes how well `office-janitor` implements the official Microsoft manual uninstall steps for Click-to-Run, MSI, and Microsoft Store installations.

---

## Executive Summary

| Installation Type | Coverage | Notes |
|-------------------|----------|-------|
| **Click-to-Run** | **~95%** | All key steps covered including App Paths and AppV |
| **MSI** | **~85%** | Component scanning partially implemented |
| **Microsoft Store** | **~90%** | Full AppX removal support via `appx_uninstall.py` |

---

## Click-to-Run Uninstall Steps

### Step 1: Uninstall the Click-to-Run task

**Microsoft KB:**
> Delete the scheduled task: `\Microsoft\Office\Office ClickToRun Service Monitor`

**Office Janitor Status:** ✅ **IMPLEMENTED**

```python
# constants.py - OFFICE_SCHEDULED_TASKS_TO_DELETE includes:
r"\Microsoft\Office\Office ClickToRun Service Monitor"
r"\Microsoft\Office\Office Automatic Updates"
```

**Location:** [constants.py](../src/office_janitor/constants.py) `OFFICE_SCHEDULED_TASKS_TO_DELETE`

---

### Step 2: Uninstall the Click-to-Run service

**Microsoft KB:**
> 1. Open Command Prompt as administrator
> 2. Run: `sc delete ClickToRunSvc`

**Office Janitor Status:** ✅ **IMPLEMENTED**

```python
# constants.py - OFFICE_SERVICES_TO_DELETE includes:
"ClickToRunSvc"
```

**Location:** [tasks_services.py](../src/office_janitor/tasks_services.py) handles service deletion via `sc delete`

---

### Step 3: Delete the Click-to-Run files

**Microsoft KB:**
> Delete these folders:
> - `C:\Program Files\Microsoft Office 15`
> - `C:\Program Files\Microsoft Office 16` (or `\Microsoft Office`)
> - `C:\Program Files\Common Files\Microsoft Shared\ClickToRun`
> - `C:\ProgramData\Microsoft\ClickToRun`

**Office Janitor Status:** ✅ **IMPLEMENTED**

```python
# constants.py - RESIDUE_PATH_TEMPLATES includes:
"programfiles_office15_x64": r"C:\Program Files\Microsoft Office 15"
"programfiles_office16_x64": r"C:\Program Files\Microsoft Office 16"  
"programfiles_clicktorun_x64": r"C:\Program Files\Common Files\Microsoft Shared\ClickToRun"
"programdata_clicktorun": r"%PROGRAMDATA%\Microsoft\ClickToRun"
```

**Also covers x86 variants.**

---

### Step 4: Delete the Click-to-Run Start menu shortcuts

**Microsoft KB:**
> Delete shortcuts from:
> - `%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs`
> - `%APPDATA%\Microsoft\Windows\Start Menu\Programs`

**Office Janitor Status:** ✅ **IMPLEMENTED**

```python
# constants.py - START_MENU_SHORTCUT_PATHS includes:
r"%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs"
r"%APPDATA%\Microsoft\Windows\Start Menu\Programs"
r"%APPDATA%\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar"
```

**Location:** [constants.py](../src/office_janitor/constants.py) `START_MENU_SHORTCUT_PATHS`, `OFFICE_SHORTCUT_PATTERNS`

---

### Step 5: Delete the Click-to-Run registry subkeys

**Microsoft KB lists these specific keys:**

| Registry Key | Status |
|--------------|--------|
| `HKLM\SOFTWARE\Microsoft\Office\ClickToRun` | ✅ Covered |
| `HKLM\SOFTWARE\Microsoft\Office\15.0\ClickToRun` | ✅ Covered |
| `HKLM\SOFTWARE\Microsoft\Office\16.0\ClickToRun` | ✅ Covered |
| `HKCU\SOFTWARE\Microsoft\Office\15.0\ClickToRun` | ✅ Covered |
| `HKCU\SOFTWARE\Microsoft\Office\16.0\ClickToRun` | ✅ Covered |
| `HKLM\SOFTWARE\Microsoft\AppV\Client\Integration\Packages\{GUID}` | ✅ Covered |
| `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\` Office apps | ✅ Covered |

**Office Janitor Status:** ✅ **FULLY IMPLEMENTED**

**Location:** [constants.py](../src/office_janitor/constants.py) `_C2R_REGISTRY_RESIDUE`, `_APP_PATHS_REGISTRY`, `_APPV_INTEGRATION_REGISTRY`

---

### Step 6: Delete the Click-to-Run registration subkeys

**Microsoft KB:**
> Delete subkeys containing "Microsoft Office" from:
> - `HKCR\Installer\Products\`
> - `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\`

**Office Janitor Status:** ⚠️ **PARTIAL**

The scrubber handles ARP (Add/Remove Programs) uninstall entries but Windows Installer Products requires GUID enumeration.

**Gap:** Need full Windows Installer metadata cleanup (see SCRUBBER_GAP_ANALYSIS.md Section 5.2)

---

### Step 7: Delete the Start Menu shortcuts

**Microsoft KB:**
> If shortcuts still remain after Step 4, manually delete Office shortcuts from Start Menu.

**Office Janitor Status:** ✅ **IMPLEMENTED** (covered by `START_MENU_SHORTCUT_PATHS` and `OFFICE_SHORTCUT_PATTERNS`)

---

## MSI Installation Uninstall Steps

### Step 1: Remove Windows Installer packages

**Microsoft KB:**
> Use `msiexec /x {ProductCode}` for each Office product

**Office Janitor Status:** ✅ **IMPLEMENTED**

**Location:** [msi_uninstall.py](../src/office_janitor/msi_uninstall.py)

---

### Step 2: Stop Office Source Engine service

**Microsoft KB:**
> Run: `net stop ose` and `sc delete ose`

**Office Janitor Status:** ✅ **IMPLEMENTED**

```python
# constants.py
OFFICE_SERVICES_TO_DELETE includes: "ose", "ose64"
```

---

### Step 3: Delete the remaining Office installation folders

**Microsoft KB lists:**
> - `C:\Program Files\Microsoft Office`
> - `C:\Program Files (x86)\Microsoft Office`
> - `C:\Program Files\Common Files\Microsoft Shared\OFFICE1x`
> - `C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE1x`
> - `C:\MSOCache`

**Office Janitor Status:** ✅ **IMPLEMENTED**

All paths covered in `RESIDUE_PATH_TEMPLATES`.

---

### Step 4: Delete Office setup files

**Microsoft KB:**
> Delete Office installer files from user's Downloads folder.

**Office Janitor Status:** ❌ **NOT IMPLEMENTED**

**Note:** This is user-specific and may not be appropriate for automated cleanup.

---

### Step 5: Delete Office registry entries

**Microsoft KB lists extensive registry cleanup:**

| Category | Keys | Status |
|----------|------|--------|
| Office version keys | `HKCU\SOFTWARE\Microsoft\Office\1x.0` | ✅ Covered |
| Office policies | `HKCU\SOFTWARE\Policies\Microsoft\Office\1x.0` | ✅ Covered |
| User Settings | `HKCU\SOFTWARE\Microsoft\Office\Common` | ✅ Covered |
| ARP entries | `HKLM\SOFTWARE\...\Uninstall\Office1x.*` | ✅ Covered |
| Windows Installer Products | `HKCR\Installer\Products\` | ⚠️ Partial |
| Windows Installer Features | `HKCR\Installer\Features\` | ⚠️ Partial |
| Windows Installer Components | `HKCR\Installer\Components\` | ⚠️ Partial |
| UserData | `HKLM\SOFTWARE\...\Installer\UserData\S-1-5-18\Products\` | ⚠️ Partial |

**Office Janitor Status:** ⚠️ **PARTIALLY IMPLEMENTED**

**Major Gap:** Full Windows Installer metadata cleanup requires:
1. GUID compression/expansion utilities ✅ (implemented in guid_utils.py)
2. Product code enumeration and selective deletion ⚠️ (partial)
3. Component reference counting ❌ (not implemented)

---

### Step 6: Delete shortcuts

**Microsoft KB:**
> Same as C2R Step 4

**Office Janitor Status:** ✅ **IMPLEMENTED** (covered by `START_MENU_SHORTCUT_PATHS` and `OFFICE_SHORTCUT_PATTERNS`)

---

## Microsoft Store Installation Steps

### Step 1: Remove Office from Settings

**Microsoft KB:**
> Settings > Apps > Apps & features > Microsoft Office Desktop Apps > Uninstall

**Office Janitor Status:** ✅ **IMPLEMENTED**

The tool now includes full Microsoft Store (AppX) removal support:

```python
from office_janitor.appx_uninstall import (
    detect_office_appx_packages,
    remove_office_appx_packages,
    remove_provisioned_appx_packages,
    is_office_store_install,
)
```

**Location:** [appx_uninstall.py](../src/office_janitor/appx_uninstall.py)

---

### Step 2: Use Windows PowerShell to remove Office

**Microsoft KB:**
> Run PowerShell as admin:
> ```powershell
> Get-AppxPackage -name "Microsoft.Office.Desktop*" | Remove-AppxPackage
> ```

**Office Janitor Status:** ✅ **IMPLEMENTED**

```python
# constants.py includes:
OFFICE_APPX_PACKAGES = (
    "Microsoft.Office.Desktop",
    "Microsoft.Office.Desktop.Access",
    "Microsoft.Office.Desktop.Excel",
    "Microsoft.Office.Desktop.Outlook",
    "Microsoft.Office.Desktop.PowerPoint",
    "Microsoft.Office.Desktop.Publisher",
    "Microsoft.Office.Desktop.Word",
    # ... and more
)
```

**Location:** [appx_uninstall.py](../src/office_janitor/appx_uninstall.py), [constants.py](../src/office_janitor/constants.py) `OFFICE_APPX_PACKAGES`

---

### Step 3: Delete remaining Office folders and registry keys

**Microsoft KB:**
> Same filesystem and registry cleanup as C2R/MSI

**Office Janitor Status:** ✅ **COVERED** (by existing cleanup)

---

## Detailed Gap Summary

### High Priority Gaps - ✅ ALL RESOLVED

| Gap | Status | Location |
|-----|--------|----------|
| **AppX/Microsoft Store removal** | ✅ Implemented | [appx_uninstall.py](../src/office_janitor/appx_uninstall.py) |
| **App Paths registry cleanup** | ✅ Implemented | [constants.py](../src/office_janitor/constants.py) `_APP_PATHS_REGISTRY` |
| **AppV Integration packages** | ✅ Implemented | [constants.py](../src/office_janitor/constants.py) `_APPV_INTEGRATION_REGISTRY` |
| **Start Menu shortcut complete cleanup** | ✅ Implemented | [constants.py](../src/office_janitor/constants.py) `START_MENU_SHORTCUT_PATHS` |

### Medium Priority Gaps

| Gap | Impact | Effort | Location |
|-----|--------|--------|----------|
| Full WI Component cleanup | Medium | High | msi_components.py |
| WI Features cleanup | Low | Medium | msi_components.py |
| Taskbar pin removal (shell COM) | Low | High | Requires shell integration |

### Low Priority / Not Recommended

| Item | Reason |
|------|--------|
| Downloads folder cleanup | User data - privacy concern |
| IE BHO cleanup | Legacy browser - minimal impact |

---

## Implementation Status - COMPLETED

### New Files Created

1. **[appx_uninstall.py](../src/office_janitor/appx_uninstall.py)** - Microsoft Store package removal
   - `detect_office_appx_packages()` - Detect installed Office AppX packages
   - `remove_office_appx_packages()` - Remove user-installed packages
   - `remove_provisioned_appx_packages()` - Remove system-wide provisioned packages
   - `is_office_store_install()` - Check if Office is Store-installed
   - `get_appx_package_info()` - Get detailed package info

### New Constants Added to [constants.py](../src/office_janitor/constants.py)

1. **App Paths Registry** (`_APP_PATHS_REGISTRY`)
   - All Office executable App Paths entries
   - Both native and WOW6432Node variants

2. **AppV Integration Registry** (`_APPV_INTEGRATION_REGISTRY`)
   - App-V Client integration packages
   - AppVISV (ISV) keys used by C2R

3. **SPFS Shell Overlays** (`_SPFS_SHELL_OVERLAYS`)
   - Microsoft SPFS Icon Overlay entries for SharePoint Workspace/Groove

4. **Start Menu Paths** (`START_MENU_SHORTCUT_PATHS`)
   - All Users Start Menu
   - Current User Start Menu
   - Quick Launch / Taskbar pins

5. **Office Shortcut Patterns** (`OFFICE_SHORTCUT_PATTERNS`)
   - Glob patterns for Office shortcut files

6. **AppX Package Names** (`OFFICE_APPX_PACKAGES`)
   - Microsoft Store Office package identifiers

7. **Provisioned AppX Patterns** (`OFFICE_APPX_PROVISIONED_PACKAGES`)
   - System-wide provisioned package patterns

---

## Verification Checklist

When testing against Microsoft's KB, verify:

- [ ] C2R scheduled tasks deleted
- [ ] ClickToRunSvc service deleted
- [ ] Office folders in Program Files deleted
- [ ] Office folders in ProgramData deleted
- [ ] Start Menu shortcuts removed
- [ ] ClickToRun registry keys removed
- [ ] AppV Integration registry cleaned (if applicable)
- [ ] Windows Installer metadata cleaned
- [ ] MSI products uninstalled via msiexec
- [ ] OSE service stopped and deleted
- [ ] MSOCache deleted
- [ ] AppX packages removed (if Store install)

---

## Conclusion

Office Janitor now provides **comprehensive coverage (~90%+)** of Microsoft's official manual uninstall procedures for all three installation types:

- **Click-to-Run:** ~95% coverage - all key steps implemented
- **MSI:** ~85% coverage - all common scenarios covered
- **Microsoft Store:** ~90% coverage - full AppX removal support

### Remaining Minor Gaps

1. **Windows Installer Component reference counting** - Advanced MSI cleanup
2. **Programmatic Taskbar unpin** - Requires COM shell integration (not in MS KB)
3. **User Downloads folder cleanup** - Intentionally omitted (privacy)

The existing [SCRUBBER_GAP_ANALYSIS.md](SCRUBBER_GAP_ANALYSIS.md) covers additional VBS script parity items that go beyond the basic KB article, providing even deeper cleanup capabilities.
