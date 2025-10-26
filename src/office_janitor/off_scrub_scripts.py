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
from pathlib import Path
from typing import Mapping


_SCRIPT_BODIES: Mapping[str, str] = {
    "OffScrub03.vbs": "' Minimal placeholder for Office 2003 MSI OffScrub\n"
    "WScript.Quit 0\n",
    "OffScrub07.vbs": "' Minimal placeholder for Office 2007 MSI OffScrub\n"
    "WScript.Quit 0\n",
    "OffScrub10.vbs": "' Minimal placeholder for Office 2010 MSI OffScrub\n"
    "WScript.Quit 0\n",
    "OffScrub_O15msi.vbs": "' Minimal placeholder for Office 2013 MSI OffScrub\n"
    "WScript.Quit 0\n",
    "OffScrub_O16msi.vbs": "' Minimal placeholder for Office 2016+ MSI OffScrub\n"
    "WScript.Quit 0\n",
    "OffScrubC2R.vbs": "' Minimal placeholder for Click-to-Run OffScrub\n"
    "WScript.Quit 0\n",
}
"""!
@brief Embedded VBS shims keyed by script file name.
@details Each entry is intentionally small: the shim merely exits successfully
to keep integration tests deterministic. Real uninstall behaviour is provided
by Python orchestration rather than the external scripts.
"""


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
        path.write_text(_SCRIPT_BODIES[script_name], encoding="utf-8")
    return path

