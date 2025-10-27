"""!
@brief Click-to-Run uninstall orchestration utilities.
@details Stops supporting services, prefers ``OfficeC2RClient.exe`` for removal,
falls back to ``setup.exe`` when necessary, and verifies registry/file system
state to confirm the Click-to-Run footprint has been removed.
"""
from __future__ import annotations

import time
from dataclasses import dataclass
from pathlib import Path
from typing import List, Mapping, MutableMapping, Sequence

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

C2R_CLIENT_CANDIDATES = (
    Path(r"C:\\Program Files\\Common Files\\Microsoft Shared\\ClickToRun\\OfficeC2RClient.exe"),
    Path(r"C:\\Program Files (x86)\\Common Files\\Microsoft Shared\\ClickToRun\\OfficeC2RClient.exe"),
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


def _collect_release_ids(raw: object) -> List[str]:
    """!
    @brief Normalise release identifier metadata into a list of strings.
    """

    if raw is None:
        return []
    if isinstance(raw, str):
        return [part.strip() for part in raw.replace(";", ",").split(",") if part.strip()]
    if isinstance(raw, Sequence):
        items: List[str] = []
        for value in raw:
            text = str(value).strip()
            if text:
                items.append(text)
        return items
    return []


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
                release_ids = [text]

    display_name = str(
        mapping.get("product")
        or property_map.get("product")
        or (", ".join(release_ids) if release_ids else "Click-to-Run Suite")
    )

    uninstall_handles: Sequence[str] = ()
    raw_handles = mapping.get("uninstall_handles")
    if isinstance(raw_handles, Sequence) and not isinstance(raw_handles, (str, bytes)):
        uninstall_handles = [str(handle).strip() for handle in raw_handles if str(handle).strip()]

    install_paths: List[Path] = []
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

    client_candidates: List[Path] = []
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

    setup_candidates: List[Path] = []
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
            human_logger.info(
                "Confirmed Click-to-Run removal for %s", target.display_name
            )
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
    total_attempts = max(1, int(retries) + 1)

    if client_path is not None:
        command = [str(client_path), *C2R_CLIENT_ARGS]
        result: command_runner.CommandResult | None = None
        for attempt in range(1, total_attempts + 1):
            message = (
                "Uninstalling Click-to-Run suite %s [attempt %d/%d]"
                % (target.display_name, attempt, total_attempts)
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
                break
            if result.returncode == 0:
                break
            if attempt < total_attempts:
                human_logger.warning(
                    "Retrying Click-to-Run uninstall via OfficeC2RClient.exe"
                )
                time.sleep(C2R_RETRY_DELAY)
        if result is not None and not (result.skipped or dry_run) and result.returncode != 0:
            raise RuntimeError("Click-to-Run uninstall failed via OfficeC2RClient.exe")
    else:
        setup_path = _find_existing_path(target.setup_candidates)
        if setup_path is None:
            raise FileNotFoundError("Neither OfficeC2RClient.exe nor setup.exe were found")
        release_ids = list(target.release_ids) or ["ALL"]
        for release_id in release_ids:
            message = f"Uninstalling Click-to-Run release {release_id} via setup.exe"
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
                },
            )
            if not dry_run and result.returncode != 0:
                raise RuntimeError(f"setup.exe uninstall failed for {release_id}")

    if dry_run:
        return

    if not _await_removal(target):
        raise RuntimeError("Click-to-Run removal verification failed")
