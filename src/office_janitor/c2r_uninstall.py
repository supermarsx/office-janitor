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
                h for h in target.uninstall_handles 
                if _parse_registry_handle(h) and registry_tools.key_exists(*_parse_registry_handle(h))
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
