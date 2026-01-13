"""!
@brief Click-to-Run uninstall orchestration utilities.
@details Stops supporting services, prefers ``OfficeC2RClient.exe`` for removal,
falls back to ``setup.exe`` when necessary, and verifies registry/file system
state to confirm the Click-to-Run footprint has been removed.
"""

from __future__ import annotations

import time
from collections.abc import Mapping, MutableMapping, Sequence
from dataclasses import dataclass
from pathlib import Path

from . import command_runner, constants, logging_ext, registry_tools, tasks_services

C2R_CLIENT_ARGS = (
    "/updatepromptuser=False",
    "/uninstallpromptuser=False",
    "/uninstall",
    "/displaylevel=False",
)
"""!
@brief Arguments passed to ``OfficeC2RClient.exe`` to trigger silent uninstall.
"""

C2R_CLIENT_FORCE_ARGS = (
    "/updatepromptuser=False",
    "/uninstallpromptuser=False",
    "/uninstall",
    "/displaylevel=False",
    "/forceappshutdown",
)
"""!
@brief Force-mode arguments for ``OfficeC2RClient.exe`` that close running Office apps.
"""

C2R_CLIENT_CANDIDATES = (
    Path(r"C:\\Program Files\\Common Files\\Microsoft Shared\\ClickToRun\\OfficeC2RClient.exe"),
    Path(
        r"C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\ClickToRun\\OfficeC2RClient.exe"
    ),
)
"""!
@brief Default filesystem locations checked for ``OfficeC2RClient.exe``.
"""

C2R_SETUP_CANDIDATES = (
    Path(r"C:\\Program Files\\Common Files\\Microsoft Shared\\ClickToRun\\setup.exe"),
    Path(r"C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\ClickToRun\\setup.exe"),
)
"""!
@brief Default filesystem locations checked for ``setup.exe`` fallback.
"""

C2R_TIMEOUT = 3600
"""!
@brief Default timeout for Click-to-Run uninstall commands.
"""

C2R_RETRY_ATTEMPTS = 1
"""!
@brief Number of retries after the initial uninstall attempt when using ``OfficeC2RClient.exe``.
"""

C2R_RETRY_DELAY = 10.0
"""!
@brief Delay between Click-to-Run uninstall retries in seconds.
"""

C2R_VERIFICATION_ATTEMPTS = 3
"""!
@brief Number of verification probes issued after uninstall commands complete.
"""

C2R_VERIFICATION_DELAY = 5.0
"""!
@brief Delay between verification probes for Click-to-Run removal.
"""

# ODT (Office Deployment Tool) URLs for download
ODT_DOWNLOAD_URLS = {
    16: "https://officecdn.microsoft.com/pr/wsus/setup.exe",
    15: "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_x86_5031-1000.exe",
}
"""!
@brief URLs for downloading ODT setup.exe by Office version.
@details Version 16 covers Office 365/2016/2019/2021/2024.
Version 15 covers Office 2013.
"""

ODT_REMOVE_XML_TEMPLATE = """<Configuration>
  <Remove All="TRUE" />
  <Display Level="{level}" AcceptEULA="TRUE" />
</Configuration>
"""
"""!
@brief Template for ODT configuration XML to remove all Office products.
"""


_C2R_RELEASE_LOOKUP = {key.lower(): key for key in constants.C2R_PRODUCT_RELEASES.keys()}
"""!
@brief Case-insensitive mapping for Click-to-Run release identifiers.
"""


C2R_RELEASE_HANDLE_CATEGORIES = {"product_release_ids"}
"""!
@brief Metadata categories that yield per-release uninstall handles.
"""


@dataclass
class _C2RTarget:
    """!
    @brief Normalised metadata for a Click-to-Run uninstall target.
    """

    display_name: str
    release_ids: Sequence[str]
    uninstall_handles: Sequence[str]
    install_paths: Sequence[Path]
    client_candidates: Sequence[Path]
    setup_candidates: Sequence[Path]


def _collect_release_ids(raw: object) -> list[str]:
    """!
    @brief Normalise release identifier metadata into a list of strings.
    """

    if raw is None:
        return []
    if isinstance(raw, str):
        return [
            _canonical_release_id(part) for part in raw.replace(";", ",").split(",") if part.strip()
        ]
    if isinstance(raw, Sequence):
        items: list[str] = []
        for value in raw:
            text = str(value).strip()
            if text:
                items.append(_canonical_release_id(text))
        return items
    return []


def _canonical_release_id(identifier: str) -> str:
    """!
    @brief Return the canonical Click-to-Run release identifier.
    """

    normalised = identifier.strip()
    if not normalised:
        return ""
    canonical = _C2R_RELEASE_LOOKUP.get(normalised.lower())
    return canonical or normalised


def _normalise_c2r_entry(config: Mapping[str, object]) -> _C2RTarget:
    """!
    @brief Convert a configuration mapping into a :class:`_C2RTarget` record.
    """

    mapping: MutableMapping[str, object] = dict(config)
    properties = mapping.get("properties")
    property_map = dict(properties) if isinstance(properties, Mapping) else {}

    release_ids = _collect_release_ids(
        mapping.get("release_ids")
        or mapping.get("products")
        or mapping.get("ProductReleaseIds")
        or property_map.get("release_id")
        or property_map.get("release_ids")
    )
    if not release_ids:
        single = mapping.get("release_id") or property_map.get("release_id")
        if single:
            text = str(single).strip()
            if text:
                release_ids = [_canonical_release_id(text)]

    release_metadata: MutableMapping[str, Mapping[str, object]] = {}
    for release_id in release_ids:
        metadata = constants.C2R_PRODUCT_RELEASES.get(_canonical_release_id(release_id))
        if metadata:
            release_metadata[release_id] = metadata

    display_name = str(
        mapping.get("product")
        or property_map.get("product")
        or next(
            (str(meta.get("product")) for meta in release_metadata.values() if meta.get("product")),
            None,
        )
        or (", ".join(release_ids) if release_ids else "Click-to-Run Suite")
    )

    uninstall_handles: Sequence[str] = ()
    raw_handles = mapping.get("uninstall_handles")
    if isinstance(raw_handles, Sequence) and not isinstance(raw_handles, (str, bytes)):
        uninstall_handles = [str(handle).strip() for handle in raw_handles if str(handle).strip()]

    derived_handles: list[str] = []
    if release_metadata:
        for release_id, metadata in release_metadata.items():
            registry_paths = metadata.get("registry_paths")
            if not isinstance(registry_paths, Mapping):
                continue
            for category, paths in registry_paths.items():
                if category not in C2R_RELEASE_HANDLE_CATEGORIES:
                    continue
                if not isinstance(paths, Sequence):
                    continue
                for hive, base in paths:
                    suffix = release_id if category == "product_release_ids" else ""
                    location = f"{base}\\{suffix}" if suffix else base
                    derived_handles.append(
                        registry_tools.hive_name(hive) + "\\" + location.strip("\\")
                    )
    if derived_handles:
        uninstall_handles = list(dict.fromkeys([*uninstall_handles, *derived_handles]))

    install_paths: list[Path] = []
    for candidate in (
        mapping.get("install_path"),
        property_map.get("install_path"),
        property_map.get("ClientFolder"),
    ):
        if not candidate:
            continue
        try:
            path = Path(str(candidate)).expanduser()
        except (TypeError, ValueError):
            continue
        install_paths.append(path)

    client_candidates: list[Path] = []
    raw_client = mapping.get("client_path") or property_map.get("client_path")
    if raw_client:
        client_candidates.append(Path(str(raw_client)))
    raw_client_paths = mapping.get("client_paths") or property_map.get("client_paths")
    if isinstance(raw_client_paths, Sequence) and not isinstance(raw_client_paths, (str, bytes)):
        for value in raw_client_paths:
            try:
                client_candidates.append(Path(str(value)))
            except (TypeError, ValueError):
                continue
    for base in install_paths:
        client_candidates.append(base / "OfficeC2RClient.exe")
    client_candidates.extend(C2R_CLIENT_CANDIDATES)

    setup_candidates: list[Path] = []
    raw_setup = mapping.get("setup_path") or property_map.get("setup_path")
    if raw_setup:
        setup_candidates.append(Path(str(raw_setup)))
    raw_setup_paths = mapping.get("setup_paths") or property_map.get("setup_paths")
    if isinstance(raw_setup_paths, Sequence) and not isinstance(raw_setup_paths, (str, bytes)):
        for value in raw_setup_paths:
            try:
                setup_candidates.append(Path(str(value)))
            except (TypeError, ValueError):
                continue
    for base in install_paths:
        setup_candidates.append(base / "setup.exe")
    setup_candidates.extend(C2R_SETUP_CANDIDATES)

    return _C2RTarget(
        display_name=display_name,
        release_ids=tuple(dict.fromkeys(release_ids)),
        uninstall_handles=tuple(uninstall_handles),
        install_paths=tuple(dict.fromkeys(install_paths)),
        client_candidates=tuple(dict.fromkeys(client_candidates)),
        setup_candidates=tuple(dict.fromkeys(setup_candidates)),
    )


def _parse_registry_handle(handle: str) -> tuple[int, str] | None:
    """!
    @brief Parse ``HKLM\\`` style handles into hive/path tuples.
    """

    cleaned = str(handle).strip()
    if not cleaned or "\\" not in cleaned:
        return None
    prefix, _, path = cleaned.partition("\\")
    hive = constants.REGISTRY_ROOTS.get(prefix.upper())
    if hive is None or not path:
        return None
    return hive, path


def _handles_present(target: _C2RTarget) -> bool:
    """!
    @brief Check whether any uninstall handles remain in the registry.
    """

    for handle in target.uninstall_handles:
        parsed = _parse_registry_handle(handle)
        if parsed and registry_tools.key_exists(parsed[0], parsed[1]):
            return True
    return False


def _install_paths_present(target: _C2RTarget) -> bool:
    """!
    @brief Check whether the recorded install paths still exist on disk.
    """

    for path in target.install_paths:
        try:
            if path.exists():
                return True
        except OSError:
            continue
    return False


def _await_removal(target: _C2RTarget) -> bool:
    """!
    @brief Poll registry and filesystem to confirm Click-to-Run removal.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    for attempt in range(1, C2R_VERIFICATION_ATTEMPTS + 1):
        registry_present = _handles_present(target)
        filesystem_present = _install_paths_present(target)
        machine_logger.info(
            "c2r_uninstall_verify",
            extra={
                "event": "c2r_uninstall_verify",
                "release_ids": list(target.release_ids) or None,
                "attempt": attempt,
                "registry_present": registry_present,
                "filesystem_present": filesystem_present,
            },
        )
        if not registry_present and not filesystem_present:
            human_logger.info("Confirmed Click-to-Run removal for %s", target.display_name)
            return True
        if attempt < C2R_VERIFICATION_ATTEMPTS:
            time.sleep(C2R_VERIFICATION_DELAY)
    return False


def _find_existing_path(candidates: Sequence[Path]) -> Path | None:
    """!
    @brief Return the first existing path from ``candidates``.
    """

    for candidate in candidates:
        try:
            if candidate.exists():
                return candidate
        except OSError:
            continue
    return None


def uninstall_products(
    config: Mapping[str, object],
    *,
    dry_run: bool = False,
    retries: int = C2R_RETRY_ATTEMPTS,
    force: bool = False,
) -> None:
    """!
    @brief Remove Click-to-Run installations using the preferred tooling.
    @details Services are stopped before invoking ``OfficeC2RClient.exe`` with
    retry semantics. If the client is unavailable the fallback ``setup.exe`` is
    invoked per release identifier. Post-uninstall verification checks registry
    handles and install paths, raising :class:`RuntimeError` when residue
    remains.
    @param config Inventory mapping describing the Click-to-Run suite.
    @param dry_run When ``True`` log planned actions without executing commands.
    @param retries Additional attempts after the first failure when using the
    client executable.
    @param force When ``True`` use force-shutdown arguments to close running Office apps.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    target = _normalise_c2r_entry(config)
    machine_logger.info(
        "c2r_uninstall_plan",
        extra={
            "event": "c2r_uninstall_plan",
            "release_ids": list(target.release_ids) or None,
            "display_name": target.display_name,
            "dry_run": bool(dry_run),
            "handles": list(target.uninstall_handles),
            "install_paths": [str(path) for path in target.install_paths],
        },
    )

    if not dry_run:
        human_logger.info("Stopping Click-to-Run related services prior to uninstall.")
        tasks_services.stop_services(constants.KNOWN_SERVICES)
    else:
        human_logger.info(
            "Dry-run: would stop Click-to-Run services %s",
            ", ".join(constants.KNOWN_SERVICES),
        )

    client_path = _find_existing_path(target.client_candidates)
    setup_path = _find_existing_path(target.setup_candidates)
    total_attempts = max(1, int(retries) + 1)

    client_succeeded = False
    client_tried = False

    if client_path is not None:
        client_tried = True
        machine_logger.info(
            "c2r_uninstall_client",
            extra={
                "event": "c2r_uninstall_client",
                "executable": str(client_path),
                "release_ids": list(target.release_ids) or None,
                "dry_run": bool(dry_run),
                "force": bool(force),
            },
        )
        # Use force args when force mode is enabled
        client_args = C2R_CLIENT_FORCE_ARGS if force else C2R_CLIENT_ARGS
        command = [str(client_path), *client_args]
        result: command_runner.CommandResult | None = None
        for attempt in range(1, total_attempts + 1):
            message = (
                f"Uninstalling Click-to-Run suite {target.display_name} "
                f"[attempt {attempt}/{total_attempts}]"
            )
            result = command_runner.run_command(
                command,
                event="c2r_uninstall",
                timeout=C2R_TIMEOUT,
                dry_run=dry_run,
                human_message=message,
                extra={
                    "release_ids": list(target.release_ids) or None,
                    "attempt": attempt,
                    "attempts": total_attempts,
                    "executable": str(client_path),
                },
            )
            if result.skipped:
                client_succeeded = True
                break
            if result.returncode == 0:
                client_succeeded = True
                break
            if attempt < total_attempts:
                human_logger.warning("Retrying Click-to-Run uninstall via OfficeC2RClient.exe")
                time.sleep(C2R_RETRY_DELAY)

        if not client_succeeded and result is not None and not (result.skipped or dry_run):
            machine_logger.warning(
                "c2r_uninstall_client_failure",
                extra={
                    "event": "c2r_uninstall_client_failure",
                    "release_ids": list(target.release_ids) or None,
                    "executable": str(client_path),
                    "return_code": result.returncode,
                },
            )
            human_logger.warning(
                "OfficeC2RClient.exe failed (exit code %d), trying setup.exe fallback...",
                result.returncode,
            )

    # Try setup.exe fallback if client wasn't found or failed
    if not client_succeeded and setup_path is not None:
        machine_logger.info(
            "c2r_uninstall_fallback",
            extra={
                "event": "c2r_uninstall_fallback",
                "release_ids": list(target.release_ids) or None,
                "executable": str(setup_path),
                "dry_run": bool(dry_run),
                "after_client_failure": client_tried,
            },
        )
        release_ids = list(target.release_ids) or ["ALL"]
        setup_failed = False
        for release_id in release_ids:
            message = f"Uninstalling Click-to-Run release {release_id} via setup.exe"
            # Add /forceappshutdown for setup.exe when in force mode
            if force:
                command = [str(setup_path), "/uninstall", release_id, "/forceappshutdown"]
            else:
                command = [str(setup_path), "/uninstall", release_id]
            result = command_runner.run_command(
                command,
                event="c2r_setup_uninstall",
                timeout=C2R_TIMEOUT,
                dry_run=dry_run,
                human_message=message,
                extra={
                    "release_id": release_id,
                    "executable": str(setup_path),
                    "force": bool(force),
                },
            )
            if not dry_run and result.returncode != 0:
                machine_logger.error(
                    "c2r_uninstall_setup_failure",
                    extra={
                        "event": "c2r_uninstall_setup_failure",
                        "release_id": release_id,
                        "executable": str(setup_path),
                        "return_code": result.returncode,
                    },
                )
                setup_failed = True

        if not setup_failed:
            client_succeeded = True  # setup.exe worked
        elif client_tried:
            raise RuntimeError(
                "Click-to-Run uninstall failed via both OfficeC2RClient.exe and setup.exe"
            )
        else:
            raise RuntimeError("Click-to-Run uninstall failed via setup.exe")

    # If neither executable was found
    if client_path is None and setup_path is None:
        raise FileNotFoundError("Neither OfficeC2RClient.exe nor setup.exe were found")

    if dry_run:
        return

    if not _await_removal(target):
        if force:
            # In force mode, try to manually clean up residue
            human_logger.warning(
                "Click-to-Run verification failed, attempting manual cleanup in force mode..."
            )
            _force_cleanup_residue(target)
            # Check again after force cleanup
            if _await_removal(target):
                human_logger.info("Force cleanup succeeded - Click-to-Run residue removed")
                return
            # Still have residue - report what remains
            remaining_paths = [str(p) for p in target.install_paths if p.exists()]
            remaining_handles = [
                h
                for h in target.uninstall_handles
                if _parse_registry_handle(h)
                and registry_tools.key_exists(*_parse_registry_handle(h))
            ]
            machine_logger.error(
                "c2r_uninstall_residue",
                extra={
                    "event": "c2r_uninstall_residue",
                    "release_ids": list(target.release_ids) or None,
                    "remaining_paths": remaining_paths,
                    "remaining_handles": remaining_handles,
                    "force_cleanup_attempted": True,
                },
            )
            raise RuntimeError(
                f"Click-to-Run removal verification failed even after force cleanup. "
                f"Remaining: {len(remaining_paths)} paths, {len(remaining_handles)} registry keys"
            )
        else:
            machine_logger.error(
                "c2r_uninstall_residue",
                extra={
                    "event": "c2r_uninstall_residue",
                    "release_ids": list(target.release_ids) or None,
                },
            )
            raise RuntimeError("Click-to-Run removal verification failed")


def _force_cleanup_residue(target: _C2RTarget) -> None:
    """!
    @brief Aggressively remove Click-to-Run residue in force mode.
    @details Attempts to delete remaining install paths and registry keys
    that the standard uninstall process failed to remove.
    """
    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    # Try to delete remaining filesystem paths
    for path in target.install_paths:
        try:
            if path.exists():
                human_logger.info("Force removing path: %s", path)
                if path.is_dir():
                    import shutil

                    shutil.rmtree(path, ignore_errors=True)
                else:
                    path.unlink(missing_ok=True)
                machine_logger.info(
                    "c2r_force_cleanup_path",
                    extra={
                        "event": "c2r_force_cleanup_path",
                        "path": str(path),
                        "success": not path.exists(),
                    },
                )
        except OSError as exc:
            human_logger.warning("Failed to force-remove path %s: %s", path, exc)

    # Try to delete remaining registry keys
    for handle in target.uninstall_handles:
        parsed = _parse_registry_handle(handle)
        if parsed and registry_tools.key_exists(parsed[0], parsed[1]):
            hive, subkey = parsed
            # Reconstruct the full registry path for delete_keys
            hive_name = registry_tools.hive_name(hive)
            full_key = f"{hive_name}\\{subkey}"
            try:
                human_logger.info("Force removing registry key: %s", full_key)
                registry_tools.delete_keys([full_key])
                machine_logger.info(
                    "c2r_force_cleanup_registry",
                    extra={
                        "event": "c2r_force_cleanup_registry",
                        "handle": handle,
                        "full_key": full_key,
                        "success": True,
                    },
                )
            except Exception as exc:
                human_logger.warning("Failed to force-remove registry key %s: %s", full_key, exc)
                machine_logger.warning(
                    "c2r_force_cleanup_registry_failed",
                    extra={
                        "event": "c2r_force_cleanup_registry_failed",
                        "handle": handle,
                        "full_key": full_key,
                        "error": str(exc),
                    },
                )


# ---------------------------------------------------------------------------
# ODT (Office Deployment Tool) Integration
# ---------------------------------------------------------------------------


def build_remove_xml(
    output_path: Path | str,
    *,
    quiet: bool = True,
) -> Path:
    """!
    @brief Build ODT configuration XML for complete Office removal.
    @details VBS equivalent: BuildRemoveXml in OffScrubC2R.vbs.
    @param output_path Path where the XML file should be written.
    @param quiet If True, use silent display level; otherwise use Full.
    @returns Path to the written XML file.
    """
    level = "None" if quiet else "Full"
    content = ODT_REMOVE_XML_TEMPLATE.format(level=level)

    path = Path(output_path)
    path.write_text(content, encoding="utf-8")

    return path


def download_odt(
    version: int = 16,
    dest_dir: Path | str | None = None,
    *,
    dry_run: bool = False,
) -> Path | None:
    """!
    @brief Download Office Deployment Tool if not available locally.
    @details VBS equivalent: DownloadODT in OffScrubC2R.vbs.
    Uses urllib to fetch ODT setup.exe from Microsoft CDN.
    @param version Office version (15 or 16). Version 16 covers 365/2016/2019/2021/2024.
    @param dest_dir Directory to save the downloaded file. Defaults to temp dir.
    @param dry_run If True, only log what would be done without downloading.
    @returns Path to downloaded setup.exe, or None if download failed.
    """
    import tempfile
    import urllib.request
    import urllib.error

    human_logger = logging_ext.get_human_logger()

    url = ODT_DOWNLOAD_URLS.get(version)
    if not url:
        human_logger.error("No ODT download URL for Office version %d", version)
        return None

    if dest_dir is None:
        dest_dir = Path(tempfile.gettempdir())
    else:
        dest_dir = Path(dest_dir)

    dest_path = dest_dir / "setup.exe"

    if dry_run:
        human_logger.info("[DRY-RUN] Would download ODT from: %s to %s", url, dest_path)
        return dest_path

    human_logger.info("Downloading ODT from: %s", url)

    try:
        # Add User-Agent to avoid potential blocks
        request = urllib.request.Request(
            url,
            headers={"User-Agent": "OfficeJanitor/1.0"},
        )
        with urllib.request.urlopen(request, timeout=60) as response:
            content = response.read()

        dest_path.write_bytes(content)
        human_logger.info("Downloaded ODT to: %s (%d bytes)", dest_path, len(content))
        return dest_path

    except urllib.error.URLError as exc:
        human_logger.error("Failed to download ODT: %s", exc)
        return None
    except OSError as exc:
        human_logger.error("Failed to save ODT: %s", exc)
        return None


def find_or_download_odt(
    version: int = 16,
    *,
    dry_run: bool = False,
) -> Path | None:
    """!
    @brief Find local ODT setup.exe or download if not found.
    @param version Office version for download URL selection.
    @param dry_run If True, don't actually download.
    @returns Path to setup.exe or None if unavailable.
    """
    human_logger = logging_ext.get_human_logger()

    # First check standard locations
    for candidate in C2R_SETUP_CANDIDATES:
        if candidate.exists():
            human_logger.debug("Found local ODT at: %s", candidate)
            return candidate

    # Check C2R client location which may have setup.exe
    for candidate in C2R_CLIENT_CANDIDATES:
        setup_path = candidate.parent / "setup.exe"
        if setup_path.exists():
            human_logger.debug("Found local ODT at: %s", setup_path)
            return setup_path

    # Not found locally - download
    human_logger.info("ODT not found locally, attempting download...")
    return download_odt(version, dry_run=dry_run)


def uninstall_via_odt(
    odt_path: Path | str,
    config_xml: Path | str | None = None,
    *,
    quiet: bool = True,
    dry_run: bool = False,
    timeout: int = C2R_TIMEOUT,
) -> int:
    """!
    @brief Execute ODT-based Office uninstall.
    @details VBS equivalent: UninstallOfficeC2R using ODT in OffScrubC2R.vbs.
    @param odt_path Path to ODT setup.exe.
    @param config_xml Path to removal XML. If None, creates a temporary one.
    @param quiet If True, use silent mode.
    @param dry_run If True, only log what would be done.
    @param timeout Command timeout in seconds.
    @returns Exit code from setup.exe (0 = success).
    """
    import tempfile

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    odt_path = Path(odt_path)
    if not odt_path.exists():
        human_logger.error("ODT setup.exe not found: %s", odt_path)
        return 1

    # Build config XML if not provided
    if config_xml is None:
        temp_dir = Path(tempfile.gettempdir())
        config_xml = build_remove_xml(temp_dir / "RemoveAll.xml", quiet=quiet)
    else:
        config_xml = Path(config_xml)

    command = [str(odt_path), "/configure", str(config_xml)]

    if dry_run:
        human_logger.info("[DRY-RUN] Would execute: %s", " ".join(command))
        return 0

    human_logger.info("Executing ODT removal: %s", " ".join(command))
    machine_logger.info(
        "odt_uninstall_start",
        extra={
            "event": "odt_uninstall_start",
            "odt_path": str(odt_path),
            "config_xml": str(config_xml),
        },
    )

    result = command_runner.run_command(
        command,
        timeout=timeout,
        event="odt_uninstall",
    )

    machine_logger.info(
        "odt_uninstall_complete",
        extra={
            "event": "odt_uninstall_complete",
            "exit_code": result.returncode,
        },
    )

    if result.returncode == 0:
        human_logger.info("ODT removal completed successfully")
    else:
        human_logger.warning("ODT removal exited with code: %d", result.returncode)

    return result.returncode


# ---------------------------------------------------------------------------
# Integrator.exe C2R unregistration
# ---------------------------------------------------------------------------

INTEGRATOR_EXE_CANDIDATES = (
    Path(r"C:\Program Files\Common Files\Microsoft Shared\ClickToRun\integrator.exe"),
    Path(r"C:\Program Files (x86)\Common Files\Microsoft Shared\ClickToRun\integrator.exe"),
)
"""!
@brief Default locations for the C2R integrator executable.
"""


def find_integrator_exe() -> Path | None:
    """!
    @brief Locate the C2R integrator.exe binary.
    @returns Path to integrator.exe or None if not found.
    """
    for candidate in INTEGRATOR_EXE_CANDIDATES:
        if candidate.exists():
            return candidate
    return None


def delete_c2r_manifests(
    package_folder: Path | str,
    *,
    dry_run: bool = False,
) -> list[Path]:
    """!
    @brief Delete C2RManifest*.xml files from a package's Integration folder.
    @details VBS equivalent: del command for C2RManifest*.xml in OffScrubC2R.vbs.
    @param package_folder Root folder of the C2R package (e.g., C:\\Program Files\\Microsoft Office).
    @param dry_run If True, only log what would be deleted.
    @returns List of manifest files deleted (or that would be deleted in dry-run).
    """
    import glob

    human_logger = logging_ext.get_human_logger()
    package_folder = Path(package_folder)

    integration_path = package_folder / "root" / "Integration"
    if not integration_path.exists():
        human_logger.debug("Integration folder not found: %s", integration_path)
        return []

    manifest_pattern = str(integration_path / "C2RManifest*.xml")
    manifest_files = [Path(p) for p in glob.glob(manifest_pattern)]

    deleted: list[Path] = []
    for manifest in manifest_files:
        if dry_run:
            human_logger.info("[DRY-RUN] Would delete manifest: %s", manifest)
            deleted.append(manifest)
        else:
            try:
                manifest.unlink()
                human_logger.debug("Deleted manifest: %s", manifest)
                deleted.append(manifest)
            except OSError as exc:
                human_logger.warning("Failed to delete manifest %s: %s", manifest, exc)

    return deleted


def unregister_c2r_integration(
    package_folder: Path | str,
    package_guid: str,
    *,
    dry_run: bool = False,
    timeout: int = 120,
) -> int:
    """!
    @brief Unregister C2R integration components via integrator.exe.
    @details VBS equivalent: integrator.exe /U /Extension call in OffScrubC2R.vbs.
        Steps:
        1. Delete C2RManifest*.xml files from the Integration folder
        2. Call integrator.exe /U /Extension with PackageRoot and PackageGUID
    @param package_folder Root folder of the C2R package.
    @param package_guid The PackageGUID for unregistration.
    @param dry_run If True, only log what would be done.
    @param timeout Timeout for integrator.exe command.
    @returns Exit code from integrator.exe (0 = success, -1 if not found).
    """
    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()
    package_folder = Path(package_folder)

    # Step 1: Delete manifest files
    deleted_manifests = delete_c2r_manifests(package_folder, dry_run=dry_run)
    if deleted_manifests:
        human_logger.info("Deleted %d C2R manifest file(s)", len(deleted_manifests))

    # Step 2: Find integrator.exe
    integrator = find_integrator_exe()
    if integrator is None:
        # Also check within the package folder
        pkg_integrator = (
            package_folder
            / "root"
            / "vfs"
            / "ProgramFilesCommonX64"
            / "Microsoft Shared"
            / "ClickToRun"
            / "integrator.exe"
        )
        if pkg_integrator.exists():
            integrator = pkg_integrator
        else:
            pkg_integrator = package_folder / "root" / "Integration" / "integrator.exe"
            if pkg_integrator.exists():
                integrator = pkg_integrator

    if integrator is None:
        human_logger.debug("Integrator.exe not found, skipping unregistration")
        return -1

    # Step 3: Build and execute unregister command
    # Format: integrator.exe /U /Extension PackageRoot=<path> PackageGUID=<guid>
    command = [
        str(integrator),
        "/U",
        "/Extension",
        f"PackageRoot={package_folder}",
        f"PackageGUID={package_guid}",
    ]

    if dry_run:
        human_logger.info("[DRY-RUN] Would execute: %s", " ".join(command))
        return 0

    human_logger.info("Unregistering C2R integration for: %s", package_folder)
    machine_logger.info(
        "c2r_unregister_start",
        extra={
            "event": "c2r_unregister_start",
            "package_folder": str(package_folder),
            "package_guid": package_guid,
        },
    )

    result = command_runner.run_command(
        command,
        timeout=timeout,
        event="c2r_unregister",
    )

    machine_logger.info(
        "c2r_unregister_complete",
        extra={
            "event": "c2r_unregister_complete",
            "exit_code": result.returncode,
        },
    )

    if result.returncode == 0:
        human_logger.debug("C2R unregistration completed successfully")
    else:
        human_logger.warning("C2R unregistration exited with code: %d", result.returncode)

    return result.returncode


def find_c2r_package_guids() -> list[tuple[Path, str]]:
    """!
    @brief Find installed C2R package folders and their GUIDs from registry.
    @details Scans HKLM\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration for
        PackageGUID and installation path information.
    @returns List of (package_folder, package_guid) tuples.
    """
    import winreg

    human_logger = logging_ext.get_human_logger()
    results: list[tuple[Path, str]] = []

    config_keys = [
        r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
        r"SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration",
    ]

    for key_path in config_keys:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_READ) as key:
                try:
                    package_guid = winreg.QueryValueEx(key, "PackageGUID")[0]
                except FileNotFoundError:
                    continue

                # Try to get install path
                install_path = None
                for value_name in ("InstallationPath", "ClientFolder"):
                    try:
                        install_path = winreg.QueryValueEx(key, value_name)[0]
                        break
                    except FileNotFoundError:
                        continue

                if install_path and package_guid:
                    results.append((Path(install_path), package_guid))
                    human_logger.debug("Found C2R package: %s (%s)", install_path, package_guid)

        except FileNotFoundError:
            continue
        except OSError as exc:
            human_logger.debug("Failed to read %s: %s", key_path, exc)
            continue

    return results


def unregister_all_c2r_integrations(*, dry_run: bool = False) -> int:
    """!
    @brief Unregister all found C2R integration components.
    @details Discovers all C2R packages from registry and unregisters each.
    @param dry_run If True, only log what would be done.
    @returns Number of packages successfully unregistered.
    """
    human_logger = logging_ext.get_human_logger()

    packages = find_c2r_package_guids()
    if not packages:
        human_logger.debug("No C2R packages found for unregistration")
        return 0

    success_count = 0
    for package_folder, package_guid in packages:
        result = unregister_c2r_integration(
            package_folder,
            package_guid,
            dry_run=dry_run,
        )
        if result == 0:
            success_count += 1

    human_logger.info("Unregistered %d of %d C2R packages", success_count, len(packages))
    return success_count


# License reinstallation support for C2R products


def get_c2r_product_release_ids() -> list[str]:
    """!
    @brief Get Office C2R product release IDs (SKUs) from registry.
    @details Scans ProductReleaseIDs under the active configuration to find
        installed Office SKUs like ProPlus, Professional, Standard, etc.
    @returns List of SKU names (e.g., ["ProPlus2024Retail", "VisioProRetail"]).
    """
    import winreg

    human_logger = logging_ext.get_human_logger()
    skus: list[str] = []

    prids_keys = [
        r"SOFTWARE\Microsoft\Office\ClickToRun\ProductReleaseIDs",
        r"SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\ProductReleaseIDs",
    ]

    for key_path in prids_keys:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_READ) as key:
                # Get ActiveConfiguration to find the channel
                try:
                    active_config = winreg.QueryValueEx(key, "ActiveConfiguration")[0]
                except FileNotFoundError:
                    continue

                # Open the configuration subkey
                config_path = f"{key_path}\\{active_config}"
                try:
                    with winreg.OpenKey(
                        winreg.HKEY_LOCAL_MACHINE, config_path, 0, winreg.KEY_READ
                    ) as config_key:
                        # Enumerate subkeys to find product SKUs
                        idx = 0
                        while True:
                            try:
                                subkey_name = winreg.EnumKey(config_key, idx)
                                # SKUs have ".16" suffix - extract just the name
                                if subkey_name.endswith(".16"):
                                    sku_name = subkey_name[:-3]  # Remove ".16"
                                    skus.append(sku_name)
                                    human_logger.debug("Found SKU: %s", sku_name)
                                idx += 1
                            except OSError:
                                break
                except FileNotFoundError:
                    continue

        except FileNotFoundError:
            continue
        except OSError as exc:
            human_logger.debug("Failed to read %s: %s", key_path, exc)
            continue

    return skus


def get_c2r_install_root() -> tuple[Path | None, str | None]:
    """!
    @brief Get C2R install root path and package GUID.
    @details Reads InstallPath and PackageGUID from registry, appending '\\root'
        to InstallPath as expected by integrator.exe /R /License.
    @returns Tuple of (install_root_path, package_guid) or (None, None) if not found.
    """
    import winreg

    human_logger = logging_ext.get_human_logger()

    config_keys = [
        r"SOFTWARE\Microsoft\Office\ClickToRun",
        r"SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun",
    ]

    for key_path in config_keys:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_READ) as key:
                try:
                    install_path = winreg.QueryValueEx(key, "InstallPath")[0]
                    package_guid = winreg.QueryValueEx(key, "PackageGUID")[0]

                    # Add \root as per OfficeScrubber.cmd logic
                    install_root = Path(install_path) / "root"
                    human_logger.debug(
                        "Found C2R install root: %s (GUID: %s)", install_root, package_guid
                    )
                    return install_root, package_guid
                except FileNotFoundError:
                    continue
        except FileNotFoundError:
            continue
        except OSError as exc:
            human_logger.debug("Failed to read %s: %s", key_path, exc)
            continue

    return None, None


def reinstall_c2r_license(
    sku_name: str,
    package_root: Path | str,
    package_guid: str,
    *,
    dry_run: bool = False,
    timeout: int = 120,
) -> int:
    """!
    @brief Reinstall Office C2R license for a single product SKU.
    @details Calls integrator.exe /R /License to reinstall license files.
        Based on OfficeScrubber.cmd license reset functionality (option T).
    @param sku_name Product SKU name (e.g., "ProPlus2024Retail").
    @param package_root C2R package root path (with \\root suffix).
    @param package_guid C2R package GUID.
    @param dry_run If True, only log what would be done.
    @param timeout Timeout for integrator.exe command.
    @returns Exit code from integrator.exe (0 = success).
    """
    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    # Find integrator.exe
    integrator = find_integrator_exe()
    if integrator is None:
        # Check within package folder
        pkg_integrator = Path(package_root) / "Integration" / "integrator.exe"
        if pkg_integrator.exists():
            integrator = pkg_integrator
        else:
            human_logger.warning("Integrator.exe not found, cannot reinstall license")
            return -1

    # Build command: integrator.exe /R /License PRIDName=<sku>.16 PackageGUID=<guid> PackageRoot=<path>
    prid_name = f"{sku_name}.16"
    command = [
        str(integrator),
        "/R",
        "/License",
        f"PRIDName={prid_name}",
        f"PackageGUID={package_guid}",
        f"PackageRoot={package_root}",
    ]

    if dry_run:
        human_logger.info("[DRY-RUN] Would reinstall license: %s", prid_name)
        return 0

    human_logger.info("Reinstalling license for: %s", sku_name)
    machine_logger.info(
        "c2r_license_reinstall_start",
        extra={
            "event": "c2r_license_reinstall_start",
            "sku_name": sku_name,
            "prid_name": prid_name,
            "package_root": str(package_root),
            "package_guid": package_guid,
        },
    )

    result = command_runner.run_command(
        command,
        timeout=timeout,
        event="c2r_license_reinstall",
    )

    machine_logger.info(
        "c2r_license_reinstall_complete",
        extra={
            "event": "c2r_license_reinstall_complete",
            "sku_name": sku_name,
            "exit_code": result.returncode,
        },
    )

    if result.returncode == 0:
        human_logger.debug("License reinstall completed for %s", sku_name)
    else:
        human_logger.warning(
            "License reinstall for %s exited with code: %d", sku_name, result.returncode
        )

    return result.returncode


def reinstall_c2r_licenses(*, dry_run: bool = False, timeout: int = 120) -> dict[str, int]:
    """!
    @brief Reinstall all Office C2R licenses using integrator.exe.
    @details Resets Office licensing by reinstalling license files for all
        detected product SKUs. Based on OfficeScrubber.cmd license menu option T.
        Steps:
        1. Detect installed C2R configuration (InstallPath, PackageGUID)
        2. Enumerate product SKUs from ProductReleaseIDs
        3. Call integrator.exe /R /License for each SKU
    @param dry_run If True, only log what would be done.
    @param timeout Timeout for each integrator.exe command.
    @returns Dictionary mapping SKU names to exit codes.
    """
    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    # Step 1: Get install root and package GUID
    install_root, package_guid = get_c2r_install_root()
    if install_root is None or package_guid is None:
        human_logger.warning("No installed Office C2R detected, cannot reinstall licenses")
        return {}

    # Check integrator.exe exists
    integrator = find_integrator_exe()
    if integrator is None:
        pkg_integrator = install_root / "Integration" / "integrator.exe"
        if not pkg_integrator.exists():
            human_logger.warning("Integrator.exe not found, cannot reinstall licenses")
            return {}

    # Step 2: Get product SKUs
    skus = get_c2r_product_release_ids()
    if not skus:
        human_logger.warning("No product SKUs found, cannot reinstall licenses")
        return {}

    human_logger.info("Found %d Office product SKU(s) for license reinstall", len(skus))
    machine_logger.info(
        "c2r_licenses_reinstall_start",
        extra={
            "event": "c2r_licenses_reinstall_start",
            "sku_count": len(skus),
            "skus": skus,
            "package_root": str(install_root),
            "package_guid": package_guid,
            "dry_run": dry_run,
        },
    )

    # Step 3: Reinstall each SKU
    results: dict[str, int] = {}
    for sku in skus:
        exit_code = reinstall_c2r_license(
            sku,
            install_root,
            package_guid,
            dry_run=dry_run,
            timeout=timeout,
        )
        results[sku] = exit_code

    successes = sum(1 for code in results.values() if code == 0)
    failures = len(results) - successes

    human_logger.info("License reinstall complete: %d succeeded, %d failed", successes, failures)
    machine_logger.info(
        "c2r_licenses_reinstall_complete",
        extra={
            "event": "c2r_licenses_reinstall_complete",
            "successes": successes,
            "failures": failures,
            "results": results,
        },
    )

    return results
