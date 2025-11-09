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
from typing import Mapping, Sequence, List

from . import constants


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

    executable = str(template.get("executable", constants.OFFSCRUB_EXECUTABLE))
    host_args = list(template.get("host_args", constants.OFFSCRUB_HOST_ARGS))

    args: List[str] = []
    script_path = None

    if kind == "msi":
        script_name = _pick_msi_script(version)
        script_path = ensure_offscrub_script(script_name, base_directory=base_directory)
        args.extend([str(part) for part in template.get("arguments", ())])
    elif kind == "c2r":
        script_name = template.get("script") or constants.C2R_OFFSCRUB_SCRIPT
        script_path = ensure_offscrub_script(script_name, base_directory=base_directory)
        args.extend([str(part) for part in template.get("arguments", ())])
    else:
        # Unknown kinds have been rejected earlier but keep safe handling
        raise KeyError(f"Unsupported OffScrub kind: {kind}")

    final: List[str] = [executable, *host_args, str(script_path), *args]
    if extra_args:
        final.extend([str(part) for part in extra_args])
    return final


def ensure_all_offscrub_shims(*, base_directory: Path | None = None) -> List[Path]:
    """!
    @brief Materialise all known OffScrub shim scripts and return their paths.
    @param base_directory Optional directory where the scripts should be written.
    @returns List of filesystem paths for each generated script.
    """

    paths: List[Path] = []
    for name in _SCRIPT_BODIES.keys():
        paths.append(ensure_offscrub_script(name, base_directory=base_directory))
    return paths


