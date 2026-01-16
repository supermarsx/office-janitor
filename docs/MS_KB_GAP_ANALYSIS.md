# Office Janitor: Microsoft KB 4e2904ea Gap Analysis

**Source:** [Manually uninstall Office (KB 4e2904ea)](https://support.microsoft.com/en-us/office/manually-uninstall-office-4e2904ea-25c8-4544-99ee-17696bb3027b)

This document analyzes how well `office-janitor` implements the official Microsoft manual uninstall steps for Click-to-Run, MSI, and Microsoft Store installations.

---

## Executive Summary

| Installation Type | Coverage | Notes |
|-------------------|----------|-------|
| **Click-to-Run** | **~85%** | Most steps covered; some shell integration gaps |
| **MSI** | **~75%** | Component scanning partially implemented |
| **Microsoft Store** | **~30%** | Requires PowerShell AppX removal; minimal support |

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

**Office Janitor Status:** ⚠️ **PARTIAL**

The scrubber has shortcut handling but may not cover all Start Menu locations consistently.

**Gap:** Need to verify shortcut cleanup covers:
- All Users Start Menu programs
- Current User Start Menu programs
- Taskbar pins (requires shell integration)

**Location:** [fs_tools.py](../src/office_janitor/fs_tools.py) - shortcut cleanup logic

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
| `HKLM\SOFTWARE\Microsoft\AppV\Client\Integration\Packages\{GUID}` | ⚠️ Partial |
| `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\` Office apps | ⚠️ Not explicitly listed |

**Office Janitor Status:** ⚠️ **MOSTLY IMPLEMENTED**

**Gap:** 
- AppV Integration packages cleanup may need GUID enumeration
- App Paths registry entries not explicitly cleaned

**Location:** [constants.py](../src/office_janitor/constants.py) `_C2R_REGISTRY_RESIDUE`

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

**Office Janitor Status:** ⚠️ **PARTIAL** (same as Step 4)

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

**Office Janitor Status:** ⚠️ **PARTIAL**

---

## Microsoft Store Installation Steps

### Step 1: Remove Office from Settings

**Microsoft KB:**
> Settings > Apps > Apps & features > Microsoft Office Desktop Apps > Uninstall

**Office Janitor Status:** ⚠️ **MINIMAL**

The tool primarily focuses on C2R and MSI. Microsoft Store (AppX) removal requires:

```powershell
Get-AppxPackage -name "Microsoft.Office.Desktop*" | Remove-AppxPackage
```

**Gap:** Need PowerShell AppX removal integration.

---

### Step 2: Use Windows PowerShell to remove Office

**Microsoft KB:**
> Run PowerShell as admin:
> ```powershell
> Get-AppxPackage -name "Microsoft.Office.Desktop*" | Remove-AppxPackage
> ```

**Office Janitor Status:** ❌ **NOT IMPLEMENTED**

**Recommendation:** Add AppX removal support:

```python
# TODO: Add to constants.py
OFFICE_APPX_PACKAGES = (
    "Microsoft.Office.Desktop",
    "Microsoft.Office.Desktop.Access",
    "Microsoft.Office.Desktop.Excel", 
    "Microsoft.Office.Desktop.Outlook",
    "Microsoft.Office.Desktop.PowerPoint",
    "Microsoft.Office.Desktop.Publisher",
    "Microsoft.Office.Desktop.Word",
)
```

---

### Step 3: Delete remaining Office folders and registry keys

**Microsoft KB:**
> Same filesystem and registry cleanup as C2R/MSI

**Office Janitor Status:** ✅ **COVERED** (by existing cleanup)

---

## Detailed Gap Summary

### High Priority Gaps

| Gap | Impact | Effort | Location |
|-----|--------|--------|----------|
| **AppX/Microsoft Store removal** | High - growing install base | Medium | New module needed |
| **App Paths registry cleanup** | Medium - orphaned entries | Low | constants.py |
| **AppV Integration packages** | Medium - C2R leftovers | Medium | registry_tools.py |
| **Start Menu shortcut complete cleanup** | Medium - visual leftovers | Low | fs_tools.py |

### Medium Priority Gaps

| Gap | Impact | Effort | Location |
|-----|--------|--------|----------|
| Full WI Component cleanup | Medium | High | msi_components.py (new) |
| WI Features cleanup | Low | Medium | msi_components.py |
| Taskbar pin removal | Low | Medium | shell integration |

### Low Priority / Not Recommended

| Item | Reason |
|------|--------|
| Downloads folder cleanup | User data - privacy concern |
| IE BHO cleanup | Legacy browser - minimal impact |

---

## Implementation Recommendations

### 1. Add Microsoft Store (AppX) Support

```python
# New file: src/office_janitor/appx_uninstall.py

def remove_office_appx_packages(*, dry_run: bool = False) -> list[str]:
    """Remove Microsoft Store Office packages via PowerShell."""
    import subprocess
    
    packages = [
        "Microsoft.Office.Desktop",
        "Microsoft.Office.Desktop.*",
    ]
    
    removed = []
    for pattern in packages:
        cmd = f'Get-AppxPackage -name "{pattern}" | Remove-AppxPackage'
        if not dry_run:
            subprocess.run(["powershell", "-Command", cmd], check=False)
        removed.append(pattern)
    
    return removed
```

### 2. Add App Paths Registry Cleanup

```python
# Add to constants.py _REGISTRY_RESIDUE_BASE

_APP_PATHS_ENTRIES = [
    (HKLM, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe"),
    (HKLM, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winword.exe"),
    (HKLM, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\powerpnt.exe"),
    (HKLM, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\outlook.exe"),
    (HKLM, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\msaccess.exe"),
    (HKLM, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\mspub.exe"),
    (HKLM, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\onenote.exe"),
    (HKLM, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\visio.exe"),
    (HKLM, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winproj.exe"),
]
```

### 3. Enhance Shortcut Cleanup

```python
# Ensure these paths are checked:
START_MENU_PATHS = [
    r"%ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs",
    r"%APPDATA%\Microsoft\Windows\Start Menu\Programs", 
    r"%APPDATA%\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar",
]
```

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

Office Janitor provides **strong coverage (~80%)** of Microsoft's official manual uninstall procedures for Click-to-Run and MSI installations. The main gaps are:

1. **Microsoft Store (AppX) support** - Critical for modern Windows 10/11
2. **Complete Start Menu/Taskbar cleanup** - Visual cleanliness
3. **App Paths registry entries** - Minor but complete cleanup

The existing `SCRUBBER_GAP_ANALYSIS.md` covers additional VBS script parity items that go beyond the basic KB article, providing even deeper cleanup capabilities.
