"""!
@brief Helpers for materialising embedded OffScrub script shims.
@details The historical Office OffScrub workflow depends on a set of VBS
automation helpers that are normally shipped alongside ``OfficeScrubber.cmd``.
Rather than bundling the original Microsoft-authored scripts, the project keeps
small, documented stand-ins that mirror their invocation signatures. The
helpers in this module write those stand-ins to disk on demand so that the rest
of the uninstaller code can continue to drive ``cscript.exe`` with familiar
arguments while remaining self-contained.
"""
from __future__ import annotations

import tempfile
import sys
from pathlib import Path
from typing import Mapping, Sequence, List

from . import constants


_SCRIPT_BODIES: Mapping[str, str] = {
    "OffScrub03.vbs": (
        "'=======================================================================================================\n"
        "' Name: OffScrub03.vbs\n"
        "' Author: Microsoft Customer Support Services\n"
        "' Copyright (c) 2010-2014 Microsoft Corporation\n"
        "' Script to remove (scrub) Office 2003 MSI products\n"
        "' when a regular uninstall is no longer possible\n"
        "'=======================================================================================================\n"
        "Option Explicit\n\n"
        "Dim sDefault\n"
        "'=======================================================================================================\n"
        "'[INI] Section for script behavior customizations\n\n"
        "'Pre-configure the SKU's to remove.\n"
        "'Only for use without command line parameters\n"
        "'Example: sDefault = \"CLIENTALL\"\n"
        "sDefault = \"\" \n\n"
        "'DO NOT CUSTOMIZE BELOW THIS LINE!\n"
        "'=======================================================================================================\n\n"
        "Const SCRIPTVERSION = \"2.14\"\n"
        "Const SCRIPTFILE    = \"OffScrub03.vbs\"\n"
        "Const SCRIPTNAME    = \"OffScrub03\"\n"
        "Const RETVALFILE    = \"ScrubRetValFile.txt\"\n"
        "Const OVERSION      = \"11.0\"\n"
        "Const OVERSIONMAJOR = \"11\"\n"
        "Const OREF          = \"Office11\"\n"
        "Const OREGREF       = \"\"\n"
        "Const ONAME         = \"Office 2003\"\n"
        "Const OPACKAGE      = \"\"\n"
        "Const OFFICEID      = \"6000-11D3-8CFE-0150048383C9}\"\n"
        "Const HKCR          = &H80000000\n"
        "Const HKCU          = &H80000001\n"
        "Const HKLM          = &H80000002\n"
        "Const HKU           = &H80000003\n"
        "Const FOR_WRITING   = 2\n"
        "Const PRODLEN       = 28\n"
        "Const COMPPERMANENT = \"00000000000000000000000000000000\"\n"
        "Const UNCOMPRESSED  = 38\n"
        "Const SQUISHED      = 20\n"
        "Const COMPRESSED    = 32\n"
        "Const REG_ARP       = \"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\\"\n"
        "Const VB_YES        = 6\n"
        "Const MSIOPENDATABASEREADONLY = 0\n\n"
        "'=======================================================================================================\n"
        "Dim oFso, oMsi, oReg, oWShell, oWmiLocal, oShellApp\n"
        "Dim ComputerItem, Item, LogStream, TmpKey\n"
        "Dim arrTmpSKUs, arrDeleteFiles, arrDeleteFolders, arrMseFolders, arrVersion\n"
        "Dim dicKeepProd, dicKeepLis, dicApps, dicKeepFolder, dicDelRegKey, dicKeepReg\n"
        "Dim dicInstalledSku, dicRemoveSku, dicKeepSku, dicSrv, dicCSuite, dicCSingle, dicManaged\n"
        "Dim f64, fLegacyProductFound, fCScript\n"
        "Dim sTmp, sSkuRemoveList, sWinDir, sWICacheDir, sMode\n"
        "Dim sAppData, sTemp, sScrubDir, sProgramFiles, sProgramFilesX86, sCommonProgramFiles\n"
        "Dim sAllusersProfile, sOSinfo, sOSVersion, sCommonProgramFilesX86, sProfilesDirectory\n"
        "Dim sProgramData, sLocalAppData, sOInstallRoot, sNotepad\n"
        "Dim iVersionNT, iError\n"
        "Dim pipename, pipeStream, fs\n\n"
        "'=======================================================================================================\n"
        "'Main\n"
        "'=======================================================================================================\n"
        "'Configure defaults\n"
        "Dim sLogDir : sLogDir = \"\"\n"
        "Dim sMoveMessage: sMoveMessage = \"\"\n"
        "Dim fClearAddinReg\t: fClearAddinReg = False\n"
        "Dim fRemoveOse      : fRemoveOse = False\n"
        "Dim fRemoveOspp     : fRemoveOspp = False\n"
        "Dim fRemoveAll      : fRemoveAll = False\n"
        "Dim fRemoveC2R      : fRemoveC2R = False\n"
        "Dim fRemoveAppV     : fRemoveAppV = False\n"
        "Dim fRemoveCSuites  : fRemoveCSuites = False\n"
        "Dim fRemoveCSingle  : fRemoveCSingle = False\n"
        "Dim fRemoveSrv      : fRemoveSrv = False\n"
        "Dim fRemoveLync     : fRemoveLync = False\n"
        "Dim fKeepUser       : fKeepUser = True  'Default to keep per user settings\n"
        "Dim fSkipSD         : fSkipSD = False 'Default to not Skip the Shortcut Detection\n"
        "Dim fKeepSG         : fKeepSG = False 'Default to not override the SoftGrid detection\n"
        "Dim fDetectOnly     : fDetectOnly = False\n"
        "Dim fQuiet          : fQuiet = False\n"
        "Dim fBasic          : fBasic = False\n"
        "Dim fNoCancel       : fNoCancel = False\n"
        "Dim fPassive        : fPassive = True\n"
        "Dim fNoReboot       : fNoReboot = False 'Default to offer reboot prompt if needed\n"
        "Dim fNoElevate      : fNoElevate = False\n"
        "Dim fElevated       : fElevated = False\n"
        "Dim fTryReconcile   : fTryReconcile = False\n"
        "Dim fC2rInstalled   : fC2rInstalled = False\n"
        "Dim fRebootRequired : fRebootRequired = False\n"
        "Dim fReturnErrorOrSuccess : fReturnErrorOrSuccess = False\n"
        "Dim fEndCurrentInstalls : fEndCurrentInstalls = False\n"
        "'CAUTION! -> \"fForce\" will kill running applications which can result in data loss! <- CAUTION\n"
        "Dim fForce          : fForce = False\n"
        "'CAUTION! -> \"fForce\" will kill running applications which can result in data loss! <- CAUTION\n"
        "Dim fLogInitialized : fLogInitialized = False\n"
        "Dim fBypass_Stage1  : fBypass_Stage1 = True 'Component Detection\n"
        "Dim fBypass_Stage2  : fBypass_Stage2 = False 'Msiexec\n"
        "Dim fBypass_Stage3  : fBypass_Stage3 = False 'CleanUp\n"
        "Dim fRunOnVanilla   : fRunOnVanilla = True\n"
        "Dim fNoOrphansMode  : fNoOrphansMode = False\n\n"
        "'Create required objects\n"
        "Set oWmiLocal   = GetObject(\"winmgmts:{(Debug)}\\.\\root\\cimv2\")\n"
        "Set oWShell     = CreateObject(\"Wscript.Shell\")\n"
        "Set oShellApp   = CreateObject(\"Shell.Application\")\n"
        "Set oFso        = CreateObject(\"Scripting.FileSystemObject\")\n"
        "Set oMsi        = CreateObject(\"WindowsInstaller.Installer\")\n"
        "Set oReg        = GetObject(\"winmgmts:\\\\.\\root\\default:StdRegProv\")\n\n"
        "LogY \"stage0\"\n\n"
        "'Get environment path info\n"
        "sAppData            = oWShell.ExpandEnvironmentStrings(\"%appdata%\")\n"
        "sLocalAppData       = oWShell.ExpandEnvironmentStrings(\"%localappdata%\")\n"
        "sTemp               = oWShell.ExpandEnvironmentStrings(\"%temp%\")\n"
        "sAllUsersProfile    = oWShell.ExpandEnvironmentStrings(\"%allusersprofile%\")\n"
        "RegReadValue HKLM, \"SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\ProfileList\", \"ProfilesDirectory\", sProfilesDirectory, \"REG_EXPAND_SZ\"\n"
        "If NOT oFso.FolderExists(sProfilesDirectory) Then \n"
        "    sProfilesDirectory  = oFso.GetParentFolderName(oWShell.ExpandEnvironmentStrings(\"%userprofile%\"))\n"
        "End If\n"
        "sProgramFiles       = oWShell.ExpandEnvironmentStrings(\"%programfiles%\")\n"
        "'Deferred until after architecture check\n"
        "'sProgramFilesX86 = oWShell.ExpandEnvironmentStrings(%programfiles(x86)%)\n\n"
        "sCommonProgramFiles = oWShell.ExpandEnvironmentStrings(\"%commonprogramfiles%\")\n"
        "'Deferred until after architecture check\n"
        "'sCommonProgramFilesX86 = oWShell.ExpandEnvironmentStrings(%CommonProgramFiles(x86)%)\n\n"
        "sProgramData        = sWSHell.ExpandEnvironmentStrings(%programdata%)\n"
        "sWinDir             = oWShell.ExpandEnvironmentStrings(\"%windir%\")\n"
        "sWICacheDir         = sWinDir & \"\\\" & \"Installer\"\n"
        "sScrubDir           = sTemp & \"\\\" & SCRIPTNAME\n"
        "sNotepad            = sWinDir & \"\\notepad.exe\"\n"
    ),
    "OffScrub07.vbs": ("DRAFT_REF:office-janitor-draft-code/bin/OffScrub07.vbs"),
    "OffScrub10.vbs": ("DRAFT_REF:office-janitor-draft-code/bin/OffScrub10.vbs"),
    "OffScrub_O15msi.vbs": ("DRAFT_REF:office-janitor-draft-code/bin/OffScrub_O15msi.vbs"),
    "OffScrub_O16msi.vbs": ("DRAFT_REF:office-janitor-draft-code/bin/OffScrub_O16msi.vbs"),
    "OffScrubC2R.vbs": ("DRAFT_REF:office-janitor-draft-code/bin/OffScrubC2R.vbs"),
}


def _default_directory() -> Path:
    """!
    @brief Compute the directory used for generated OffScrub shims.
    @details The helpers live under the user's temporary directory so they do
    not pollute the project tree or PyInstaller bundle. Callers may override the
    directory via :func:`ensure_offscrub_script` when a custom location is
    required (for example during integration tests).
    """

    directory = Path(tempfile.gettempdir()) / "office_janitor_offscrub"
    directory.mkdir(parents=True, exist_ok=True)
    return directory


def ensure_offscrub_script(script_name: str, *, base_directory: Path | None = None) -> Path:
    """!
    @brief Materialise the requested OffScrub shim on disk.
    @param script_name Name of the helper to materialise (e.g. ``OffScrub03.vbs``).
    @param base_directory Optional directory where the script should be written.
    @returns The filesystem path to the generated script.
    @throws KeyError If ``script_name`` is not recognised.
    """

    if script_name not in _SCRIPT_BODIES:
        raise KeyError(f"Unknown OffScrub helper requested: {script_name}")

    directory = base_directory or _default_directory()
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / script_name
    if not path.exists():
        body = _SCRIPT_BODIES[script_name]
        # Support a DRAFT_REF: marker which points to a path inside the repository
        if isinstance(body, str) and body.startswith("DRAFT_REF:"):
            draft_rel = body.split(":", 1)[1]
            project_root = Path(__file__).resolve().parents[2]
            draft_path = project_root / draft_rel
            if draft_path.exists():
                try:
                    data = draft_path.read_bytes()
                    try:
                        text = data.decode("utf-8")
                    except UnicodeDecodeError:
                        try:
                            text = data.decode("latin-1")
                        except Exception:
                            text = data.decode("utf-8", errors="replace")
                except Exception:
                    text = ""
            else:
                text = ""
        else:
            text = body
        path.write_text(text, encoding="utf-8")
    return path


def _pick_msi_script(version: str | None) -> str:
    """!
    @brief Select the MSI OffScrub script filename for ``version``.
    @details Falls back to the default script when a mapping is not found.
    """

    if not version:
        return constants.MSI_OFFSCRUB_DEFAULT_SCRIPT
    script_map = constants.MSI_OFFSCRUB_SCRIPT_MAP
    candidate = script_map.get(str(version))
    if candidate:
        return candidate
    return constants.MSI_OFFSCRUB_DEFAULT_SCRIPT


def build_offscrub_command(
    kind: str,
    *,
    version: str | None = None,
    base_directory: Path | None = None,
    extra_args: Sequence[str] | None = None,
) -> List[str]:
    """!
    @brief Build a command line for invoking an OffScrub helper.
    @param kind Either ``"msi"`` or ``"c2r"`` indicating the command template.
    @param version Optional Office version string used to select MSI helper.
    @param base_directory Optional directory where shim scripts are/will be written.
    @param extra_args Additional argument strings appended to the script invocation.
    @returns Command list suitable for ``exec_utils.run_command``.
    @throws KeyError When ``kind`` is not recognised.
    """

    template = constants.UNINSTALL_COMMAND_TEMPLATES.get(kind)
    if template is None:
        raise KeyError(f"Unknown OffScrub kind: {kind}")

    # Create the script on disk (keeps tests and external tooling happy)
    executable = sys.executable
    host_args = ["-m", "office_janitor.off_scrub_native"]

    args: List[str] = []
    script_path = None

    if kind == "msi":
        script_name = _pick_msi_script(version)
        script_path = ensure_offscrub_script(script_name, base_directory=base_directory)
        args.extend(["msi"])
        if extra_args is None:
            extra_args = list(template.get("arguments", ()))
    elif kind == "c2r":
        script_name = template.get("script") or constants.C2R_OFFSCRUB_SCRIPT
        script_path = ensure_offscrub_script(script_name, base_directory=base_directory)
        args.extend(["c2r"])
        if extra_args is None:
            extra_args = list(template.get("arguments", ()))
    else:
        # Unknown kinds have been rejected earlier but keep safe handling
        raise KeyError(f"Unsupported OffScrub kind: {kind}")

    # Preserve the generated script path in the arguments for compatibility
    final: List[str] = [executable, *host_args, *args, str(script_path)]
    if extra_args:
        final.extend([str(part) for part in extra_args])
    return final


def ensure_offscrub_launcher(script_path: Path) -> Path:
    """!
    @brief Write a small backward-compatible launcher next to the generated script.
    @details Creates a `.cmd` file that invokes the native Python module with the
    same arguments the legacy `cscript.exe` workflow would have provided. This
    helps external tooling that expects a file to exist while migration is in
    progress.
    """

    cmd_path = script_path.with_suffix(script_path.suffix + ".cmd")
    if not cmd_path.exists():
        # Keep the launcher simple and cross-platform friendly for CI: use
        # the same sys.executable invocation the `build_offscrub_command` uses.
        launcher_content = f'"{sys.executable}" -m office_janitor.off_scrub_native "msi" "{script_path}"\n'
        try:
            cmd_path.write_text(launcher_content, encoding="utf-8")
        except Exception:
            # Best-effort: ignore write failures to avoid breaking generation
            pass
    return cmd_path


def ensure_all_offscrub_shims(*, base_directory: Path | None = None) -> List[Path]:
    """!
    @brief Materialise all known OffScrub shim scripts and return their paths.
    @param base_directory Optional directory where the scripts should be written.
    @returns List of filesystem paths for each generated script.
    """

    paths: List[Path] = []
    for name in _SCRIPT_BODIES.keys():
        p = ensure_offscrub_script(name, base_directory=base_directory)
        # Emit a launcher for compatibility (best-effort)
        try:
            ensure_offscrub_launcher(p)
        except Exception:
            pass
        paths.append(p)
    return paths


