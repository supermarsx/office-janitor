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

import sys
import tempfile
from collections.abc import Mapping, Sequence
from pathlib import Path

from . import constants

_SCRIPT_BODIES: Mapping[str, str] = {
    "OffScrub03.vbs": (
        "' OffScrub03.vbs compatibility shim\n"
        "' Native Python flow: python -m office_janitor.off_scrub_native msi OffScrub03.vbs ALL\n"
    ),
    "OffScrub07.vbs": (
        "' OffScrub07.vbs compatibility shim\n"
        "' Native Python flow: python -m office_janitor.off_scrub_native msi OffScrub07.vbs ALL\n"
    ),
    "OffScrub10.vbs": (
        "' OffScrub10.vbs compatibility shim\n"
        "' Native Python flow: python -m office_janitor.off_scrub_native msi OffScrub10.vbs ALL\n"
    ),
    "OffScrub_O15msi.vbs": (
        "' OffScrub_O15msi.vbs compatibility shim\n"
        "' Native Python flow: python -m office_janitor.off_scrub_native msi "
        "OffScrub_O15msi.vbs ALL\n"
    ),
    "OffScrub_O16msi.vbs": (
        "' OffScrub_O16msi.vbs compatibility shim\n"
        "' Native Python flow: python -m office_janitor.off_scrub_native msi "
        "OffScrub_O16msi.vbs ALL\n"
    ),
    "OffScrubC2R.vbs": (
        "' OffScrubC2R.vbs compatibility shim\n"
        "' Native Python flow: python -m office_janitor.off_scrub_native c2r OffScrubC2R.vbs ALL\n"
    ),
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
        path.write_text(body, encoding="utf-8")
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
) -> list[str]:
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

    args: list[str] = []
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
    final: list[str] = [executable, *host_args, *args, str(script_path)]
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
        launcher_content = (
            f'"{sys.executable}" -m office_janitor.off_scrub_native "msi" "{script_path}"\n'
        )
        try:
            cmd_path.write_text(launcher_content, encoding="utf-8")
        except Exception:
            # Best-effort: ignore write failures to avoid breaking generation
            pass
    return cmd_path


def ensure_all_offscrub_shims(*, base_directory: Path | None = None) -> list[Path]:
    """!
    @brief Materialise all known OffScrub shim scripts and return their paths.
    @param base_directory Optional directory where the scripts should be written.
    @returns List of filesystem paths for each generated script.
    """

    paths: list[Path] = []
    for name in _SCRIPT_BODIES.keys():
        p = ensure_offscrub_script(name, base_directory=base_directory)
        # Emit a launcher for compatibility (best-effort)
        try:
            ensure_offscrub_launcher(p)
        except Exception:
            pass
        paths.append(p)
    return paths
