# Office Janitor: Legacy VBS Scrubber Gap Analysis & Implementation TODO

## Executive Summary

This document provides a comprehensive gap analysis between the legacy Microsoft OffScrub VBScript tools (`OffScrub03.vbs`, `OffScrub07.vbs`, `OffScrub10.vbs`, `OffScrub_O15msi.vbs`, `OffScrub_O16msi.vbs`, `OffScrubC2R.vbs`) and the Python `office_janitor` implementation. While significant progress has been made, there are critical features from the VBS scripts that require implementation to achieve full feature parity.

---

## Table of Contents

1. [Implementation Status Overview](#1-implementation-status-overview)
2. [MSI Scrubber Gaps (OffScrub03/07/10/O15/O16)](#2-msi-scrubber-gaps)
3. [Click-to-Run Scrubber Gaps (OffScrubC2R)](#3-click-to-run-scrubber-gaps)
4. [Common Infrastructure Gaps](#4-common-infrastructure-gaps)
5. [Registry Cleanup Gaps](#5-registry-cleanup-gaps)
6. [File System Cleanup Gaps](#6-file-system-cleanup-gaps)
7. [Service & Process Management Gaps](#7-service--process-management-gaps)
8. [Windows Installer (MSI) API Gaps](#8-windows-installer-msi-api-gaps)
9. [Licensing Cleanup Gaps](#9-licensing-cleanup-gaps)
10. [Priority Implementation Roadmap](#10-priority-implementation-roadmap)
11. [Test Coverage Requirements](#11-test-coverage-requirements)

---

## 1. Implementation Status Overview

### Currently Implemented ✅

| Feature | Python Module | Coverage |
|---------|--------------|----------|
| Basic C2R detection | `detect.py` | ~70% |
| Basic MSI detection | `detect.py` | ~60% |
| C2R uninstall via OfficeC2RClient | `c2r_uninstall.py` | ~80% |
| MSI uninstall via msiexec | `msi_uninstall.py` | ~70% |
| Legacy argument parsing | `off_scrub_helpers.py` | ~75% |
| Basic registry cleanup | `registry_tools.py` | ~50% |
| Basic file cleanup | `fs_tools.py` | ~40% |
| Scheduled task deletion | `tasks_services.py` | ~60% |
| Logging (human + JSONL) | `logging_ext.py` | ~90% |
| Elevation handling | `elevation.py` | ~85% |

### Partially Implemented ⚠️

| Feature | Status | Gap |
|---------|--------|-----|
| Component scanning | Minimal | Missing full WI component enumeration |
| SKU/Product categorization | Basic | Missing suite/single/server classification |
| Shortcut unpinning | Stub only | No actual taskbar/start menu unpinning |
| MSOCache (LIS) cleanup | Partial | Missing targeted cleanup per-SKU |
| OSPP/SPP license cleanup | Partial | Missing SoftwareLicensingProduct WMI calls |
| Setup.exe removal path | Basic | Missing maintenance mode orchestration |
| TypeLib cleanup | None | Not implemented |

### Not Implemented ❌

| Feature | VBS Script | Priority |
|---------|------------|----------|
| Full Windows Installer component scanning | All MSI scripts | **CRITICAL** |
| MSI .msi file caching for detection | All MSI scripts | HIGH |
| Product GUID squishing/expansion | All scripts | HIGH |
| WI metadata validation | All scripts | HIGH |
| Orphaned file detection via ComponentPath | MSI scripts | HIGH |
| Setup.exe based uninstall | O15/O16/O07 | MEDIUM |
| Explorer.exe shell integration cleanup | All scripts | MEDIUM |
| ODT (Office Deployment Tool) download & invoke | OffScrubC2R | MEDIUM |
| C2R integrator.exe invocation | OffScrubC2R | MEDIUM |
| Reboot orchestration | All scripts | LOW |
| Named pipe progress reporting | All scripts | LOW |

---

## 2. MSI Scrubber Gaps

### 2.1 Product Detection & Classification

**VBS Implementation (OffScrub_O15msi.vbs lines 520-650):**

The VBS scripts implement sophisticated product detection that:
- Scans `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall` for Office entries
- Parses `PackageRefs`/`PackageIds` multi-string values
- Categorizes products into: Client Suites, Client Single, Server, C2R
- Handles "orphaned" products by scanning the Windows Installer cache
- Creates temporary ARP entries for standalone products

**Python Gap:**

```python
# TODO: Implement in detect.py or new msi_detect.py

def classify_msi_product(product_code: str) -> str:
    """
    Classify MSI product by its GUID structure.
    
    Product codes use this pattern: {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}
    - Byte 11-14 indicates product type
    - Byte 4-5 indicates Office major version
    
    Categories:
    - "client_suite": 000F, 0011-001B, 0029, 002B, etc.
    - "client_single": Standalone apps (Word, Excel, etc.)
    - "server": Products where byte 11 = '1'
    - "c2r": Integration components (007E, 008F, etc.)
    """
    pass

def scan_orphaned_msi_products(wi_cache_dir: str) -> list[dict]:
    """
    Scan Windows Installer cache for .msi files that have no ARP entry.
    VBS equivalent: fTryReconcile logic in FindInstalledOProducts
    """
    pass

def create_temp_arp_entry(product_code: str, product_info: dict) -> None:
    """
    Create temporary ARP entry for standalone products.
    VBS equivalent: arrTmpSKUs population in FindInstalledOProducts
    """
    pass
```

### 2.2 Component Scanning (ScanComponents)

**VBS Implementation (lines 818-1050 in OffScrub_O15msi.vbs):**

This is the most complex and critical feature missing. The VBS scripts:

1. Enumerate all Windows Installer components via `oMsi.Components`
2. For each component, enumerate clients via `oMsi.ComponentClients(ComponentID)`
3. Track which products own which components
4. Use `oMsi.ComponentPath` to get file/registry paths
5. Build deletion lists for files and registry keys
6. Handle "permanent" components (COMPPERMANENT GUID)
7. Output `FileList.txt`, `RegList.txt`, `CompVerbose.txt` logs

**Python TODO:**

```python
# TODO: Create new module src/office_janitor/msi_components.py

import win32com.client  # or ctypes for msi.dll

class MSIComponentScanner:
    """
    Scan Windows Installer component database for Office-related entries.
    
    VBS equivalent: ScanComponents subroutine
    """
    
    def __init__(self, target_product_codes: set[str]):
        self.target_products = target_product_codes
        self.delete_files: list[str] = []
        self.delete_registry: list[tuple[int, str]] = []
        self.delete_components: list[str] = []
    
    def scan(self) -> None:
        """
        Full component scan matching VBS ScanComponents.
        
        Steps:
        1. Get total component count
        2. For each component:
           a. Check all ComponentClients
           b. If all clients are in removal scope, mark for deletion
           c. Track ComponentPath (file or registry) for cleanup
        3. Handle permanent components (GUID all zeros)
        4. Generate deletion lists
        """
        pass
    
    def get_component_path(self, product_code: str, component_id: str) -> str:
        """Get the keypath for a component (file path or registry key)."""
        pass
    
    def is_permanent_component(self, component_id: str) -> bool:
        """Check if component has permanent flag (all zeros GUID)."""
        pass
```

### 2.3 Setup.exe Based Removal

**VBS Implementation (SetupExeRemoval in O07/O10/O15/O16):**

The scripts attempt removal via Office's setup.exe before falling back to msiexec:

```vbscript
' Locate setup.exe from InstallSource registry
' Build command: setup.exe /uninstall <ProductID> /config <config.xml>
' Handle NOLOCALCAB scenarios
```

**Python TODO:**

```python
# TODO: Add to msi_uninstall.py

def attempt_setup_exe_removal(product: dict, *, dry_run: bool) -> bool:
    """
    Try Office setup.exe for cleaner uninstall before msiexec fallback.
    
    Steps:
    1. Locate setup.exe from InstallSource or InstallLocation registry
    2. Build uninstall config XML
    3. Execute: setup.exe /uninstall <SKU> /config <xml>
    4. Return True if successful, False to try msiexec
    """
    pass
```

### 2.4 GUID Compression/Expansion

**VBS Implementation (GetCompressedGuid, GetExpandedGuid):**

Windows Installer stores GUIDs in "compressed" and "squished" formats in registry. The VBS scripts include conversion utilities.

**Python TODO:**

```python
# TODO: Add to registry_tools.py or new guid_utils.py

def compress_guid(guid: str) -> str:
    """
    Convert {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX} to compressed form.
    Used in: Installer\Components, Installer\Products registry
    
    Algorithm:
    1. Remove braces and dashes
    2. Reverse each segment according to WI rules
    """
    pass

def expand_guid(compressed: str) -> str:
    """Reverse of compress_guid."""
    pass

def squish_guid(guid: str) -> str:
    """Convert GUID to 20-char "squished" format for some WI APIs."""
    pass

def decode_squished_guid(squished: str) -> str:
    """Reverse of squish_guid."""
    pass
```

---

## 3. Click-to-Run Scrubber Gaps

### 3.1 ODT (Office Deployment Tool) Integration

**VBS Implementation (UninstallOfficeC2R in OffScrubC2R.vbs lines 1280-1400):**

The script:
1. Builds a `RemoveAll.xml` configuration file
2. Looks for local ODT setup.exe
3. If not found, downloads from Microsoft CDN
4. Executes: `setup.exe /configure RemoveAll.xml`

**Python TODO:**

```python
# TODO: Add to c2r_uninstall.py

def build_remove_xml(output_path: str, *, quiet: bool = True) -> str:
    """
    Build ODT configuration XML for removal.
    
    Content:
    <Configuration>
      <Remove All="TRUE" />
      <Display Level="None" />  <!-- if quiet -->
    </Configuration>
    """
    pass

def download_odt(version: int, dest_dir: str) -> str | None:
    """
    Download Office Deployment Tool if not available locally.
    
    URLs:
    - v15: https://download.microsoft.com/download/.../officedeploymenttool_x86_5031-1000.exe
    - v16: http://officecdn.microsoft.com/pr/wsus/setup.exe
    """
    pass

def uninstall_via_odt(odt_path: str, config_xml: str, *, dry_run: bool) -> int:
    """Execute ODT-based uninstall."""
    pass
```

### 3.2 Integrator.exe Invocation

**VBS Implementation (lines 1200-1250):**

```vbscript
' Delete manifest files
sCmd = "cmd.exe /c del " & Chr(34) & sPkgFld & "\root\Integration\C2RManifest*.xml" & Chr(34)
' Unregister via integrator
sCmd = integrator.exe /U /Extension PackageRoot=... PackageGUID=...
```

**Python TODO:**

```python
# TODO: Add to c2r_uninstall.py

def unregister_c2r_integration(package_folder: str, package_guid: str, *, dry_run: bool) -> None:
    """
    Unregister C2R integration components.
    
    Steps:
    1. Delete C2RManifest*.xml files
    2. Call integrator.exe /U with PackageRoot and PackageGUID
    """
    pass
```

### 3.3 OSPP License Cleanup (CleanOSPP)

**VBS Implementation (OffScrubC2R.vbs lines 1000-1050):**

```vbscript
Const OfficeAppId = "0ff1ce15-a989-479d-af46-f275c6370663"

' Query WMI for licenses
Set oProductInstances = oWmiLocal.ExecQuery(
    "SELECT ID, ApplicationId, PartialProductKey, Name, ProductKeyID " &
    "FROM SoftwareLicensingProduct " &
    "WHERE ApplicationId = '" & OfficeAppId & "' AND PartialProductKey <> NULL"
)

' Uninstall each license
For Each pi in oProductInstances
    pi.UninstallProductKey(pi.ProductKeyID)
Next
```

**Python TODO:**

```python
# TODO: Add to licensing.py

OFFICE_APPLICATION_ID = "0ff1ce15-a989-479d-af46-f275c6370663"

def clean_ospp_licenses(*, dry_run: bool = False) -> list[str]:
    """
    Remove Office licenses from Software Protection Platform.
    
    Uses WMI:
    - SoftwareLicensingProduct (Win8+)
    - OfficeSoftwareProtectionProduct (Win7)
    """
    import wmi
    
    removed = []
    c = wmi.WMI()
    
    # Try modern SPP first
    for product in c.SoftwareLicensingProduct(
        ApplicationId=OFFICE_APPLICATION_ID
    ):
        if product.PartialProductKey:
            if not dry_run:
                product.UninstallProductKey(product.ProductKeyID)
            removed.append(product.Name)
    
    return removed
```

### 3.4 vNext License Cache Cleanup

**VBS Implementation (ClearVNextLicCache):**

```vbscript
sLocalAppData = oWShell.ExpandEnvironmentStrings("%localappdata%")
DeleteFolder sLocalAppData & "\Microsoft\Office\Licenses"
```

**Python Status:** Partially covered in `constants.RESIDUE_PATH_TEMPLATES` but needs explicit handling in `off_scrub_helpers.py`.

---

## 4. Common Infrastructure Gaps

### 4.1 Windows Installer Metadata Validation

**VBS Implementation (EnsureValidWIMetadata):**

```vbscript
Sub EnsureValidWIMetadata(hDefKey, sKey, iValidLength)
    ' Ensures only valid metadata entries exist to avoid API failures
    ' Invalid entries (wrong GUID length) are removed
    If RegEnumKey(hDefKey, sKey, arrKeys) Then
        For Each SubKey in arrKeys
            If NOT Len(SubKey) = iValidLength Then
                RegDeleteKey hDefKey, sKey & "\" & SubKey & "\"
            End If
        Next
    End If
End Sub
```

**Python TODO:**

```python
# TODO: Add to registry_tools.py

def validate_wi_metadata(hive: int, key: str, expected_length: int, *, dry_run: bool) -> int:
    """
    Clean invalid Windows Installer metadata entries.
    
    Invalid entries can cause WI API failures. Removes subkeys that
    don't match expected GUID length (32 for compressed, 38 for standard).
    
    Args:
        hive: Registry hive constant
        key: Base key path
        expected_length: Expected subkey name length
        dry_run: If True, only log what would be deleted
    
    Returns:
        Number of invalid entries removed
    """
    pass
```

### 4.2 Return Code Bitmask System

**VBS Implementation (Error Constants):**

```vbscript
Const ERROR_SUCCESS                 = 0    ' Bit #1
Const ERROR_FAIL                    = 1    ' Bit #1
Const ERROR_REBOOT_REQUIRED         = 2    ' Bit #2
Const ERROR_USERCANCEL              = 4    ' Bit #3
Const ERROR_STAGE1                  = 8    ' Bit #4 - Msiexec failed
Const ERROR_STAGE2                  = 16   ' Bit #5 - Cleanup failed
Const ERROR_INCOMPLETE              = 32   ' Bit #6 - Pending renames
Const ERROR_DCAF_FAILURE            = 64   ' Bit #7 - Second attempt failed
Const ERROR_ELEVATION_USERDECLINED  = 128  ' Bit #8
Const ERROR_ELEVATION               = 256  ' Bit #9
Const ERROR_SCRIPTINIT              = 512  ' Bit #10
Const ERROR_RELAUNCH                = 1024 ' Bit #11
Const ERROR_UNKNOWN                 = 2048 ' Bit #12
```

**Python TODO:**

```python
# TODO: Add to constants.py and use throughout

from enum import IntFlag

class ScrubErrorCode(IntFlag):
    SUCCESS = 0
    FAIL = 1
    REBOOT_REQUIRED = 2
    USER_CANCEL = 4
    MSIEXEC_FAILED = 8
    CLEANUP_FAILED = 16
    INCOMPLETE = 32
    RETRY_FAILED = 64
    ELEVATION_DECLINED = 128
    ELEVATION_ERROR = 256
    INIT_ERROR = 512
    RELAUNCH_ERROR = 1024
    UNKNOWN = 2048
```

### 4.3 Named Pipe Progress Reporting (LogY)

**VBS Implementation:**

```vbscript
Sub LogY(sText)
    ' Send progress stage info through named pipe
    If Not pipename = "" Then
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set pipeStream = fs.OpenTextFile(pipename, 2, False, 0)
        pipeStream.WriteLine sText
        pipeStream.Close
    End If
End Sub

' Usage throughout:
LogY "stage0"
LogY "stage1"
LogY "CleanOSPP"
LogY "reboot"
LogY "ok"
```

**Python TODO:**

```python
# TODO: Add to logging_ext.py

class NamedPipeReporter:
    """Report progress via named pipe for external monitoring."""
    
    def __init__(self, pipe_name: str | None = None):
        self.pipe_name = pipe_name
        self._pipe_handle = None
    
    def report(self, stage: str) -> None:
        """Send stage update through named pipe."""
        if not self.pipe_name:
            return
        # Implementation for Windows named pipes
        pass
```

---

## 5. Registry Cleanup Gaps

### 5.1 Full RegWipe Implementation

**VBS Implementation (RegWipe in OffScrubC2R.vbs lines 1660-1900):**

The VBS cleanup covers dozens of registry locations:

```vbscript
' Registration keys
RegDeleteKey HKCU, "Software\Microsoft\Office\15.0\Registration"
RegDeleteKey HKCU, "Software\Microsoft\Office\16.0\Registration"

' Virtual InstallRoot
RegDeleteKey HKLM, "SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot\Virtual"
RegDeleteKey HKLM, "SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot\Virtual"

' MAPI Search registration
RegDeleteKey HKLM, "SOFTWARE\Classes\CLSID\{2027FC3B-CF9D-4ec7-A823-38BA308625CC}"

' C2R keys
RegDeleteKey HKLM, "SOFTWARE\Microsoft\Office\ClickToRun"
RegDeleteKey HKLM, "SOFTWARE\Microsoft\Office\ClickToRunStore"

' AppV ISV keys
' Run keys
' UpgradeCodes
' Global Components
' Published Components
```

**Python TODO:**

```python
# TODO: Expand constants.py REGISTRY_RESIDUE_PATHS

C2R_REGISTRY_CLEANUP = [
    # Registration
    (HKCU, r"Software\Microsoft\Office\15.0\Registration"),
    (HKCU, r"Software\Microsoft\Office\16.0\Registration"),
    (HKCU, r"Software\Microsoft\Office\Registration"),
    
    # InstallRoot Virtual
    (HKLM, r"SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot\Virtual"),
    (HKLM, r"SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot\Virtual"),
    (HKLM, r"SOFTWARE\Microsoft\Office\Common\InstallRoot\Virtual"),
    
    # C2R keys
    (HKCU, r"SOFTWARE\Microsoft\Office\15.0\ClickToRun"),
    (HKLM, r"SOFTWARE\Microsoft\Office\15.0\ClickToRun"),
    (HKLM, r"SOFTWARE\Microsoft\Office\15.0\ClickToRunStore"),
    (HKCU, r"SOFTWARE\Microsoft\Office\16.0\ClickToRun"),
    (HKLM, r"SOFTWARE\Microsoft\Office\16.0\ClickToRun"),
    (HKLM, r"SOFTWARE\Microsoft\Office\16.0\ClickToRunStore"),
    (HKCU, r"SOFTWARE\Microsoft\Office\ClickToRun"),
    (HKLM, r"SOFTWARE\Microsoft\Office\ClickToRun"),
    (HKLM, r"SOFTWARE\Microsoft\Office\ClickToRunStore"),
    
    # AppV ISV
    (HKCU, r"SOFTWARE\Microsoft\AppV\ISV"),
    (HKLM, r"SOFTWARE\Microsoft\AppV\ISV"),
    (HKCU, r"SOFTWARE\Microsoft\AppVISV"),
    (HKLM, r"SOFTWARE\Microsoft\AppVISV"),
]
```

### 5.2 Windows Installer Metadata Cleanup

**VBS Implementation (lines 1750-1850):**

```vbscript
' UpgradeCodes
sSubKeyName = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes\"
' ... enumerate and selectively delete based on product scope

' UserData Products
sSubKeyName = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\"

' Features
sSubKeyName = "Installer\Features\"

' Products
sSubKeyName = "Installer\Products\"

' Components in Global
sSubKeyName = "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components\"

' Published Components
sSubKeyName = "Installer\Components\"
```

**Python TODO:**

```python
# TODO: Add to off_scrub_helpers.py or new wi_cleanup.py

def cleanup_wi_upgrade_codes(target_codes: set[str], *, dry_run: bool) -> int:
    """Clean UpgradeCodes registry entries for removed products."""
    pass

def cleanup_wi_products(target_codes: set[str], *, dry_run: bool) -> int:
    """Clean Windows Installer Products registry."""
    pass

def cleanup_wi_features(target_codes: set[str], *, dry_run: bool) -> int:
    """Clean Windows Installer Features registry."""
    pass

def cleanup_wi_components(target_codes: set[str], *, dry_run: bool) -> int:
    """
    Clean Windows Installer Components registry.
    
    Most complex - must handle:
    - Global components (S-1-5-18)
    - Published components (HKCR\Installer\Components)
    - Multi-string values with multiple product references
    """
    pass
```

### 5.3 Shell Integration Cleanup (ClearShellIntegrationReg)

**VBS Implementation (lines 1920-2020):**

```vbscript
' Protocol Handlers
RegDeleteKey HKLM, "SOFTWARE\Classes\Protocols\Handler\osf"

' Context Menu Handlers
RegDeleteKey HKLM, "SOFTWARE\Classes\CLSID\{573FFD05-2805-47C2-BCE0-5F19512BEB8D}"
' ... more CLSIDs

' Groove ShellIconOverlayIdentifiers
RegDeleteKey HKLM, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 1 (ErrorConflict)"
' ... more overlays

' Shell extensions (Approved)
RegDeleteValue HKLM, "...\Shell Extensions\Approved\", "{CLSID}", False

' Browser Helper Objects
RegDeleteKey HKLM, "...\Browser Helper Objects\{CLSID}"

' OneNote Namespace Extension
RegDeleteKey HKLM, "...\Desktop\NameSpace\{CLSID}"
```

**Python TODO:**

```python
# TODO: Add to constants.py

SHELL_INTEGRATION_CLEANUP = {
    "protocol_handlers": [
        (HKLM, r"SOFTWARE\Classes\Protocols\Handler\osf"),
    ],
    "context_menu_clsids": [
        "{573FFD05-2805-47C2-BCE0-5F19512BEB8D}",
        "{8BA85C75-763B-4103-94EB-9470F12FE0F7}",
        "{CD55129A-B1A1-438E-A425-CEBC7DC684EE}",
        "{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}",
        "{E768CD3B-BDDC-436D-9C13-E1B39CA257B1}",
    ],
    "shell_icon_overlays": [
        "Microsoft SPFS Icon Overlay 1 (ErrorConflict)",
        "Microsoft SPFS Icon Overlay 2 (SyncInProgress)",
        "Microsoft SPFS Icon Overlay 3 (InSync)",
    ],
    "browser_helper_objects": [
        "{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}",
        "{B4F3A835-0E21-4959-BA22-42B3008E02FF}",
        "{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}",
    ],
}
```

### 5.4 TypeLib Cleanup (RegWipeTypeLib)

**VBS Implementation (lines 2025-2150):**

```vbscript
' List of known Office TypeLib GUIDs
sTypeLibs = "{000204EF-...};{00020802-...};..." ' ~80 GUIDs

' For each typelib:
'   1. Check if registered
'   2. Get file path from Win32 or Win64 subkey
'   3. If file doesn't exist, delete registration
```

**Python TODO:**

```python
# TODO: Add to constants.py and registry_tools.py

OFFICE_TYPELIBS = [
    "{000204EF-0000-0000-C000-000000000046}",  # VBA
    "{00020802-0000-0000-C000-000000000046}",  # Excel
    "{00020813-0000-0000-C000-000000000046}",  # Excel
    "{00020905-0000-0000-C000-000000000046}",  # Word
    "{0002123C-0000-0000-C000-000000000046}",  # Word
    # ... ~80 more
]

def cleanup_orphaned_typelibs(*, dry_run: bool) -> int:
    """
    Remove TypeLib registrations where target DLL/TLB no longer exists.
    
    For each known Office TypeLib GUID:
    1. Check HKLM\Software\Classes\TypeLib\{GUID}\{version}
    2. Read Win32 and Win64 subkeys for file path
    3. If file doesn't exist, remove the registration
    """
    pass
```

---

## 6. File System Cleanup Gaps

### 6.1 Full FileWipe Implementation (C2R)

**VBS Implementation (FileWipe in OffScrubC2R.vbs lines 2150-2400):**

```vbscript
' Delete C2R package files
DeleteFolder sProgramFiles & "\Microsoft Office 15"
DeleteFolder sProgramFiles & "\Microsoft Office 16"
DeleteFolder sProgramFiles & "\Microsoft Office\PackageManifests"
DeleteFolder sProgramFiles & "\Microsoft Office\PackageSunrisePolicies"
DeleteFolder sProgramFiles & "\Microsoft Office\root"
DeleteFile sProgramFiles & "\Microsoft Office\AppXManifest.xml"
DeleteFile sProgramFiles & "\Microsoft Office\FileSystemMetadata.xml"

' Common Files cleanup
DeleteFolder sCommonProgramFiles & "\Microsoft Shared\ClickToRun"
DeleteFolder sCommonProgramFiles & "\Microsoft Shared\OFFICE15"
DeleteFolder sCommonProgramFiles & "\Microsoft Shared\OFFICE16"

' User data cleanup
DeleteFolder sAppData & "\Microsoft\Office\Recent"
DeleteFolder sLocalAppData & "\Microsoft\Office\15.0"
DeleteFolder sLocalAppData & "\Microsoft\Office\16.0"

' ProgramData cleanup  
DeleteFolder sProgramData & "\Microsoft\ClickToRun"
DeleteFolder sProgramData & "\Microsoft\Office\ClickToRunPackageLocker"
```

**Python Status:** Partially covered in `constants.RESIDUE_PATH_TEMPLATES` but needs expansion.

### 6.2 MSOCache (Local Installation Source) Cleanup

**VBS Implementation (WipeLIS):**

```vbscript
' MSOCache contains local installation source files
' Delete only folders for products being removed

For Each MsoFolder in arrLIS
    If InScope(MsoFolder.Name) Then
        DeleteFolder MsoFolder.Path
    End If
Next
```

**Python TODO:**

```python
# TODO: Add to fs_tools.py

def cleanup_mso_cache(target_products: set[str], *, dry_run: bool) -> int:
    """
    Clean Local Installation Source (MSOCache) for removed products.
    
    MSOCache location: All fixed drives root + MSOCache
    Scoped by product code patterns in folder names.
    """
    pass
```

### 6.3 Windows Installer Cache Orphan Cleanup

**VBS Implementation (MsiClearOrphanedFiles):**

```vbscript
' Scan %windir%\Installer for .msi files
' Check if ProductCode is still registered
' Delete orphaned files
```

**Python TODO:**

```python
# TODO: Add to fs_tools.py

def cleanup_wi_cache_orphans(*, dry_run: bool) -> int:
    """
    Remove orphaned .msi and .msp files from Windows Installer cache.
    
    Files in %WINDIR%\Installer that don't correspond to any
    registered Windows Installer product.
    """
    pass
```

### 6.4 Shortcut Unpinning

**VBS Implementation (CleanShortcuts with unpin logic):**

```vbscript
' Uses Shell.Application verbs to unpin shortcuts
' Handles both taskbar and start menu pinning
Sub Unpin(sShortcutPath)
    Set oShell = CreateObject("Shell.Application")
    Set oFolder = oShell.Namespace(oFso.GetParentFolderName(sShortcutPath))
    Set oItem = oFolder.ParseName(oFso.GetFileName(sShortcutPath))
    For Each verb in oItem.Verbs
        If InStr(verb.Name, "Unpin") > 0 Or InStr(verb.Name, "Un&pin") > 0 Then
            verb.DoIt
            Exit For
        End If
    Next
End Sub
```

**Python TODO:**

```python
# TODO: Add to fs_tools.py

def unpin_shortcut(shortcut_path: str) -> bool:
    """
    Unpin a shortcut from taskbar/start menu using Shell verbs.
    
    Uses Windows Shell API via COM:
    1. Get Shell.Application
    2. Navigate to shortcut folder
    3. Find "Unpin" verb and execute
    """
    import win32com.client
    
    shell = win32com.client.Dispatch("Shell.Application")
    folder = shell.Namespace(os.path.dirname(shortcut_path))
    item = folder.ParseName(os.path.basename(shortcut_path))
    
    for verb in item.Verbs():
        if "unpin" in verb.Name.lower():
            verb.DoIt()
            return True
    return False
```

---

## 7. Service & Process Management Gaps

### 7.1 Service Deletion

**VBS Implementation (DeleteService):**

```vbscript
Sub DeleteService(sServiceName)
    ' Stop service
    oWShell.Run "net stop " & sServiceName, 0, True
    ' Delete service
    oWShell.Run "sc delete " & sServiceName, 0, True
End Sub

' Services deleted:
DeleteService "OfficeSvc"
DeleteService "ClickToRunSvc"
```

**Python TODO:**

```python
# TODO: Expand tasks_services.py

def delete_service(service_name: str, *, dry_run: bool) -> bool:
    """
    Stop and delete a Windows service.
    
    Steps:
    1. net stop <service>
    2. sc delete <service>
    """
    pass

OFFICE_SERVICES_TO_DELETE = [
    "OfficeSvc",
    "ClickToRunSvc",
    "ose",  # Office Source Engine
    "ospp",  # Office Software Protection Platform (if no other Office)
]
```

### 7.2 Scheduled Task Deletion

**VBS Implementation (DelSchtasks):**

```vbscript
' Many scheduled tasks to delete:
oWShell.Run "SCHTASKS /Delete /TN FF_INTEGRATEDstreamSchedule /F"
oWShell.Run "SCHTASKS /Delete /TN FF_INTEGRATEDUPDATEDETECTION /F"
oWShell.Run "SCHTASKS /Delete /TN C2RAppVLoggingStart /F"
oWShell.Run "SCHTASKS /Delete /TN " & Chr(34) & "Office 15 Subscription Heartbeat" & Chr(34) & " /F"
oWShell.Run "SCHTASKS /Delete /TN " & Chr(34) & "\Microsoft\Office\OfficeInventoryAgentFallBack" & Chr(34) & " /F"
' ... many more
```

**Python Status:** Partially implemented in `constants.C2R_CLEANUP_TASKS`. Need to verify completeness.

### 7.3 Process Termination

**VBS Implementation (CloseOfficeApps):**

```vbscript
' Known processes to terminate
dicApps.Add "appvshnotify.exe"
dicApps.Add "integratedoffice.exe"
dicApps.Add "integrator.exe"
dicApps.Add "firstrun.exe"
dicApps.Add "communicator.exe"
dicApps.Add "msosync.exe"
dicApps.Add "OneNoteM.exe"
dicApps.Add "iexplore.exe"  ' Yes, IE
dicApps.Add "officeclicktorun.exe"
dicApps.Add "officeondemand.exe"
dicApps.Add "OfficeC2RClient.exe"

' Also terminate any process with ExecutablePath in C2R folders
For Each Process in Processes
    If IsC2R(Process.ExecutablePath) Then
        Process.Terminate()
    End If
Next
```

**Python TODO:**

```python
# TODO: Expand processes.py OFFICE_PROCESSES list

OFFICE_PROCESSES_EXTENDED = [
    # Standard Office apps (already have most)
    "winword.exe", "excel.exe", "powerpnt.exe", "outlook.exe",
    "msaccess.exe", "onenote.exe", "mspub.exe", "visio.exe",
    "winproj.exe", "lync.exe", "teams.exe",
    
    # C2R infrastructure
    "appvshnotify.exe",
    "integratedoffice.exe", 
    "integrator.exe",
    "firstrun.exe",
    "officeclicktorun.exe",
    "officeondemand.exe",
    "OfficeC2RClient.exe",
    
    # Background services
    "msosync.exe",
    "OneNoteM.exe",
    "communicator.exe",
    "perfboost.exe",
    "roamingoffice.exe",
    "mavinject32.exe",
    
    # Legacy
    "bcssync.exe",
    "officesas.exe",
    "officesasscheduler.exe",
]
```

### 7.4 Explorer.exe Restart

**VBS Implementation (RestoreExplorer):**

```vbscript
Sub RestoreExplorer
    ' Check if explorer is running
    Set Processes = oWmiLocal.ExecQuery("Select * From Win32_Process Where Name = 'explorer.exe'")
    If Processes.Count = 0 Then
        ' Restart explorer
        oWShell.Run "explorer.exe", 1, False
    End If
End Sub
```

**Python TODO:**

```python
# TODO: Add to processes.py

def restart_explorer_if_needed() -> bool:
    """
    Restart explorer.exe if it's not running.
    
    Called after shell integration cleanup that may have
    terminated explorer to release file locks.
    """
    pass
```

---

## 8. Windows Installer (MSI) API Gaps

### 8.1 Full MSI API Wrapper

**VBS Implementation uses:**

```vbscript
Set oMsi = CreateObject("WindowsInstaller.Installer")

' Used APIs:
oMsi.Products                        ' Enumerate all products
oMsi.Components                      ' Enumerate all components
oMsi.ComponentClients(ComponentID)   ' Get products owning component
oMsi.ComponentPath(ProductCode, ComponentID)  ' Get keypath
oMsi.ProductInfo(ProductCode, property)       ' Get product info
```

**Python TODO:**

```python
# TODO: Create new module src/office_janitor/msi_api.py

class WindowsInstaller:
    """
    Python wrapper for Windows Installer API.
    
    Uses either:
    - win32com.client for COM automation
    - ctypes for direct msi.dll calls
    """
    
    def __init__(self):
        import win32com.client
        self._msi = win32com.client.Dispatch("WindowsInstaller.Installer")
    
    def enumerate_products(self) -> list[str]:
        """Enumerate all installed product codes."""
        return list(self._msi.Products)
    
    def enumerate_components(self) -> list[str]:
        """Enumerate all registered component GUIDs."""
        return list(self._msi.Components)
    
    def get_component_clients(self, component_id: str) -> list[str]:
        """Get product codes that own a component."""
        return list(self._msi.ComponentClients(component_id))
    
    def get_component_path(self, product_code: str, component_id: str) -> str:
        """Get the keypath (file or registry) for a component."""
        return self._msi.ComponentPath(product_code, component_id)
    
    def get_product_info(self, product_code: str, property_name: str) -> str:
        """Get a product property value."""
        return self._msi.ProductInfo(product_code, property_name)
```

---

## 9. Licensing Cleanup Gaps

### 9.1 Full License Cleanup

**VBS Implementation covers:**

1. **OSPP (Office Software Protection Platform)** - via WMI
2. **vNext License Cache** - filesystem
3. **Token-based activation files** - filesystem
4. **OSPP service** - service deletion

**Python TODO:**

```python
# TODO: Expand licensing.py

def full_license_cleanup(*, dry_run: bool, keep_license: bool = False) -> dict:
    """
    Complete Office license cleanup.
    
    Components:
    1. OSPP license removal via WMI
    2. vNext license cache deletion (%LOCALAPPDATA%\Microsoft\Office\Licenses)
    3. Token activation files
    4. Shared Computer Licensing cache
    
    Returns:
        Dict with counts of cleaned items per category
    """
    if keep_license:
        return {"skipped": True, "reason": "keep_license flag set"}
    
    results = {}
    results["ospp"] = clean_ospp_licenses(dry_run=dry_run)
    results["vnext_cache"] = clean_vnext_cache(dry_run=dry_run)
    results["tokens"] = clean_activation_tokens(dry_run=dry_run)
    results["scl_cache"] = clean_scl_cache(dry_run=dry_run)
    return results
```

---

## 10. Priority Implementation Roadmap

### Phase 1: Critical (Must Have for Parity) 🔴

| Item | Module | Effort | Description |
|------|--------|--------|-------------|
| MSI Component Scanner | `msi_components.py` | HIGH | Full WI component enumeration and tracking |
| GUID Compression/Expansion | `registry_tools.py` | MEDIUM | Required for WI metadata cleanup |
| WI Metadata Validation | `registry_tools.py` | LOW | Prevent API failures |
| Full Registry Cleanup | `off_scrub_helpers.py` | MEDIUM | All VBS RegWipe locations |
| Shell Integration Cleanup | `registry_tools.py` | MEDIUM | Protocol handlers, BHOs, overlays |

### Phase 2: High Priority (Improved Reliability) 🟠

| Item | Module | Effort | Description |
|------|--------|--------|-------------|
| OSPP License Cleanup | `licensing.py` | MEDIUM | WMI-based license removal |
| TypeLib Cleanup | `registry_tools.py` | MEDIUM | Orphaned typelib detection |
| ODT Integration | `c2r_uninstall.py` | MEDIUM | Download and invoke ODT |
| Full File Cleanup | `fs_tools.py` | LOW | Expand path lists |
| MSOCache Cleanup | `fs_tools.py` | LOW | LIS cleanup per-product |

### Phase 3: Medium Priority (Feature Complete) 🟡

| Item | Module | Effort | Description |
|------|--------|--------|-------------|
| Setup.exe Uninstall | `msi_uninstall.py` | MEDIUM | Maintenance mode removal |
| Shortcut Unpinning | `fs_tools.py` | MEDIUM | Shell verbs for unpin |
| Integrator.exe | `c2r_uninstall.py` | LOW | C2R component unregistration |
| WI Cache Orphan Cleanup | `fs_tools.py` | LOW | Orphaned .msi cleanup |
| Service Deletion | `tasks_services.py` | LOW | Stop and delete services |

### Phase 4: Low Priority (Nice to Have) 🟢

| Item | Module | Effort | Description |
|------|--------|--------|-------------|
| Named Pipe Progress | `logging_ext.py` | LOW | For external monitoring |
| Full Error Bitmask | `constants.py` | LOW | VBS-compatible return codes |
| Explorer Restart | `processes.py` | LOW | Auto-restart after shell cleanup |
| Temp ARP Entry Creation | `msi_detect.py` | MEDIUM | For orphan handling |

---

## 11. Test Coverage Requirements

### New Test Files Needed

```
tests/
  test_msi_components.py      # Component scanning tests
  test_guid_utils.py          # GUID compression tests
  test_wi_cleanup.py          # WI metadata cleanup tests
  test_shell_cleanup.py       # Shell integration tests
  test_typelib_cleanup.py     # TypeLib cleanup tests
  test_ospp_cleanup.py        # License cleanup tests
  test_mso_cache.py           # MSOCache cleanup tests
```

### Mock/Stub Requirements

1. **Windows Installer COM Object** - Mock for `WindowsInstaller.Installer`
2. **WMI SoftwareLicensingProduct** - Mock for license queries
3. **Shell.Application** - Mock for shortcut unpinning
4. **Registry State** - Comprehensive fixtures for WI metadata

### Integration Test Scenarios

Based on VBS script test scenarios:

1. **Clean C2R installation** - O365/M365 full removal
2. **Clean MSI installation** - Office 2016 MSI removal
3. **Mixed installation** - C2R + MSI on same machine
4. **Orphaned products** - Products with broken ARP entries
5. **Partial removal** - Selective SKU removal (e.g., Visio only)
6. **Dry run verification** - All paths execute but no mutations

---

## Appendix A: VBS File Line Count Reference

| Script | Total Lines | Key Sections |
|--------|-------------|--------------|
| OffScrub03.vbs | 4,236 | Office 2003 MSI |
| OffScrub07.vbs | 4,446 | Office 2007 MSI |
| OffScrub10.vbs | 4,820 | Office 2010 MSI |
| OffScrub_O15msi.vbs | 4,515 | Office 2013 MSI |
| OffScrub_O16msi.vbs | 4,606 | Office 2016 MSI |
| OffScrubC2R.vbs | 3,827 | Click-to-Run |
| **Total** | **26,450** | |

---

## Appendix B: Constants to Add

```python
# Add to constants.py

# Office Product Code Patterns
MSI_PRODUCT_TYPE_CODES = {
    # Client Suites
    "000F": "ProPlus",
    "0011": "Professional", 
    "0012": "Standard",
    "0013": "Basic",
    "0014": "Professional",
    "0015": "Access",
    "0016": "Excel",
    "0017": "SharePoint Designer",
    "0018": "PowerPoint",
    "0019": "Publisher",
    "001A": "Outlook",
    "001B": "Word",
    "0029": "Excel",
    "002B": "Word",
    # ... many more
    
    # Integration/Licensing
    "007E": "Licensing",
    "008F": "Licensing", 
    "008C": "Extensibility",
    "00DD": "Extensibility x64",
    "24E1": "MSOID Login",
    "237A": "MSOID Login",
}

# Lync/Skype Product GUIDs
LYNC_PRODUCT_GUIDS = [
    "{4A2C120F-307B-4400-B239-F29ADB54D3C6}",
    "{5CFD6599-10E5-4CF0-B6E1-BF39D30A64F8}",
    # ... from LYNC_ALL constant in VBS
]
```

---

*Document generated: January 2026*
*Last updated: Auto-generated from VBS analysis*
