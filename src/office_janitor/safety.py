"""!
@brief Safety and guardrail enforcement helpers.
@details Centralises preflight checks shared across the application, including
administrative validation, supported operating system checks, process
preflight policies, dry-run guardrails, and the existing plan validation
constraints. The functions exported here are invoked by the CLI entry point
and by destructive subsystems prior to mutating the host so that ``--dry-run``
and ``--force`` semantics remain consistent.
"""
from __future__ import annotations

from typing import Iterable, Mapping, Sequence

from . import constants, fs_tools

FILESYSTEM_WHITELIST = fs_tools.FILESYSTEM_WHITELIST

FILESYSTEM_BLACKLIST = fs_tools.FILESYSTEM_BLACKLIST

SUPPORTED_SYSTEMS = {"windows", "nt"}

MINIMUM_SUPPORTED_WINDOWS_RELEASE = (6, 1)

REGISTRY_WHITELIST = (
    r"HKLM\\SOFTWARE\\MICROSOFT\\OFFICE",
    r"HKLM\\SOFTWARE\\WOW6432NODE\\MICROSOFT\\OFFICE",
    r"HKLM\\SOFTWARE\\POLICIES\\MICROSOFT\\OFFICE",
    r"HKLM\\SOFTWARE\\WOW6432NODE\\POLICIES\\MICROSOFT\\OFFICE",
    r"HKLM\\SOFTWARE\\POLICIES\\MICROSOFT\\CLOUD\\OFFICE",
    r"HKLM\\SOFTWARE\\WOW6432NODE\\POLICIES\\MICROSOFT\\CLOUD\\OFFICE",
    r"HKLM\\SOFTWARE\\MICROSOFT\\CLICKTORUN",
    r"HKLM\\SOFTWARE\\MICROSOFT\\OFFICESOFTWAREPROTECTIONPLATFORM",
    r"HKLM\\SOFTWARE\\MICROSOFT\\WINDOWS NT\\CURRENTVERSION\\SOFTWAREPROTECTIONPLATFORM",
    r"HKCU\\SOFTWARE\\MICROSOFT\\OFFICE",
    r"HKCU\\SOFTWARE\\POLICIES\\MICROSOFT\\OFFICE",
    r"HKCU\\SOFTWARE\\POLICIES\\MICROSOFT\\CLOUD\\OFFICE",
    r"HKU\\S-1-5-20\\SOFTWARE\\MICROSOFT\\OFFICESOFTWAREPROTECTIONPLATFORM",
    r"HKU\\S-1-5-20\\SOFTWARE\\MICROSOFT\\WINDOWS NT\\CURRENTVERSION\\SOFTWAREPROTECTIONPLATFORM",
)

REGISTRY_BLACKLIST = (
    r"HKLM\\SOFTWARE\\MICROSOFT\\WINDOWS",
    r"HKCU\\SOFTWARE\\MICROSOFT\\WINDOWS",
)


def perform_preflight_checks(plan: Iterable[Mapping[str, object]]) -> None:
    """!
    @brief Validate that the plan satisfies safety requirements before execution.
    """
    plan_steps = list(plan)
    if not plan_steps:
        return

    context = _extract_context(plan_steps)
    metadata = context.get("metadata", {})
    dry_run = bool(metadata.get("dry_run", False))
    mode = str(metadata.get("mode", ""))
    targets = [str(item) for item in metadata.get("target_versions", [])]
    unsupported = metadata.get("unsupported_targets", []) or []

    if unsupported:
        raise ValueError("Unsupported Office versions requested: " + ", ".join(sorted({str(u) for u in unsupported})))

    if mode == "diagnose":
        _ensure_no_action_steps(plan_steps)
        return

    if mode == "cleanup-only":
        _ensure_no_uninstall_steps(plan_steps)

    if mode.startswith("target:"):
        _ensure_targeted_uninstalls_present(plan_steps)

    _ensure_dry_run_consistency(plan_steps, dry_run)
    _enforce_target_scope(plan_steps, targets)
    _enforce_filesystem_whitelist(plan_steps)
    _enforce_registry_whitelist(plan_steps)
    _enforce_template_guard(plan_steps, metadata)


def evaluate_runtime_environment(
    *,
    is_admin: bool,
    os_system: str,
    os_release: str,
    blocking_processes: Sequence[str],
    dry_run: bool,
    require_restore_point: bool,
    restore_point_available: bool,
    force: bool = False,
    allow_unsupported_windows: bool = False,
) -> None:
    """!
    @brief Validate runtime prerequisites before destructive execution.
    @details Aggregates guard conditions enforced before mutating the host. The
    function is intentionally side-effect free so callers can execute it during
    planning or immediately prior to scrub execution. ``--force`` only skips
    advisory guards (unsupported OS releases, lingering processes, or missing
    restore points) and never bypasses immutable requirements such as
    administrative rights.
    @param is_admin Indicates whether the current process is elevated.
    @param os_system Result of ``platform.system()`` (or equivalent).
    @param os_release Operating system release string (e.g. ``"10.0.19045"``).
    @param blocking_processes Processes that must be terminated prior to
    modification.
    @param dry_run When ``True`` destructive actions are not executed.
    @param require_restore_point Caller intent for creating a restore point.
    @param restore_point_available Whether restore points are currently
    available on the host.
    @param force Indicates ``--force`` was supplied, allowing certain guard
    rails to be bypassed.
    @param allow_unsupported_windows Overrides the Windows release minimum
    guard without relaxing other safety rails.
    @raises PermissionError If administrative rights are missing for a
    destructive run.
    @raises RuntimeError If any advisory guard blocks execution.
    """

    _enforce_admin_guard(is_admin=is_admin, dry_run=dry_run)
    _enforce_os_guard(
        os_system=os_system,
        os_release=os_release,
        force=force,
        allow_unsupported_windows=allow_unsupported_windows,
    )
    _enforce_process_guard(
        blocking_processes=blocking_processes, dry_run=dry_run, force=force
    )
    _enforce_restore_point_guard(
        require_restore_point=require_restore_point,
        restore_point_available=restore_point_available,
        dry_run=dry_run,
        force=force,
    )


def guard_destructive_action(action: str, *, dry_run: bool, force: bool = False) -> None:
    """!
    @brief Prevent destructive actions from running during dry-run simulations.
    @details Callers should invoke this helper before deleting files, removing
    registry keys, or performing any irreversible change. The guard raises a
    :class:`RuntimeError` when ``dry_run`` is active unless ``--force`` is also
    provided, mirroring the specification's expectation that destructive
    behaviour remains opt-in during simulations.
    @param action Human readable description of the operation used in the error
    message.
    @param dry_run Indicates whether the global run is dry-run only.
    @param force When ``True`` the guard is bypassed.
    @raises RuntimeError If ``dry_run`` is active without ``force``.
    """

    if dry_run and not force:
        description = action or "destructive operation"
        raise RuntimeError(
            f"{description} is blocked while running in dry-run mode. Use --force to override."
        )


def _extract_context(plan_steps: Sequence[Mapping[str, object]]) -> Mapping[str, object]:
    for step in plan_steps:
        if step.get("category") == "context":
            return step
    raise ValueError("Plan is missing context metadata.")


def _ensure_no_action_steps(plan_steps: Sequence[Mapping[str, object]]) -> None:
    for step in plan_steps:
        if step.get("category") != "context":
            raise ValueError("Diagnostics mode must not schedule action steps.")


def _ensure_no_uninstall_steps(plan_steps: Sequence[Mapping[str, object]]) -> None:
    for step in plan_steps:
        if step.get("category") in {"msi-uninstall", "c2r-uninstall"}:
            raise ValueError("Cleanup-only mode cannot include uninstall steps.")


def _ensure_targeted_uninstalls_present(plan_steps: Sequence[Mapping[str, object]]) -> None:
    for step in plan_steps:
        if step.get("category") in {"msi-uninstall", "c2r-uninstall"}:
            return
    raise ValueError("Targeted scrub requested but no matching installations were selected.")


def _ensure_dry_run_consistency(plan_steps: Sequence[Mapping[str, object]], dry_run: bool) -> None:
    for step in plan_steps:
        metadata = step.get("metadata", {})
        if "dry_run" in metadata and bool(metadata["dry_run"]) != dry_run:
            raise ValueError("Plan step dry-run flag disagrees with global selection.")


def _enforce_target_scope(plan_steps: Sequence[Mapping[str, object]], targets: Sequence[str]) -> None:
    if not targets:
        return
    allowed = {str(target) for target in targets}
    for step in plan_steps:
        if step.get("category") in {"msi-uninstall", "c2r-uninstall"}:
            version = str(step.get("metadata", {}).get("version", ""))
            if version and version not in allowed:
                raise ValueError("Plan step does not align with the selected target versions.")


def _enforce_filesystem_whitelist(plan_steps: Sequence[Mapping[str, object]]) -> None:
    for step in plan_steps:
        if step.get("category") != "filesystem-cleanup":
            continue
        metadata = step.get("metadata", {})
        for path in metadata.get("paths", []) or []:
            if not _path_allowed(path):
                raise ValueError(f"Refusing to operate on non-whitelisted path: {path}")


def _enforce_registry_whitelist(plan_steps: Sequence[Mapping[str, object]]) -> None:
    for step in plan_steps:
        if step.get("category") != "registry-cleanup":
            continue
        metadata = step.get("metadata", {})
        keys = metadata.get("keys") or metadata.get("paths") or []
        for key in keys:
            if not _registry_allowed(key):
                raise ValueError(f"Refusing to modify non-whitelisted registry key: {key}")


def _path_allowed(path: str) -> bool:
    if fs_tools.is_path_whitelisted(path, whitelist=FILESYSTEM_WHITELIST, blacklist=FILESYSTEM_BLACKLIST):
        return True
    normalized = fs_tools.normalize_windows_path(path)
    if any(normalized.startswith(fs_tools.normalize_windows_path(blocked)) for blocked in FILESYSTEM_BLACKLIST):
        return False
    return False


def _registry_allowed(key: str) -> bool:
    normalized = key.upper()
    if any(normalized.startswith(blocked) for blocked in REGISTRY_BLACKLIST):
        return False
    return any(normalized.startswith(allowed) for allowed in REGISTRY_WHITELIST)


def _enforce_template_guard(
    plan_steps: Sequence[Mapping[str, object]],
    context_metadata: Mapping[str, object],
) -> None:
    """!
    @brief Prevent user template purges without explicit consent.
    @details Mirrors the OffScrub requirement that Normal.dotm and similar user
    content remain intact unless ``--force`` or an equivalent override is
    supplied. Plan steps may still opt-in by setting ``purge_templates``.
    """

    options = context_metadata.get("options", {}) if context_metadata else {}
    global_force = bool(options.get("force", False))
    keep_templates = bool(options.get("keep_templates", False))

    for step in plan_steps:
        if step.get("category") != "filesystem-cleanup":
            continue
        metadata = step.get("metadata", {}) or {}
        preserve_flag = bool(metadata.get("preserve_templates", keep_templates))
        purge_flag = bool(metadata.get("purge_templates", False))
        paths = metadata.get("paths", []) or []

        template_targets = [
            str(path)
            for path in paths
            if _is_template_path(str(path))
        ]
        if not template_targets:
            continue

        if preserve_flag and not global_force:
            raise ValueError(
                "Plan requests user template preservation but includes template cleanup steps."
            )

        if not purge_flag and not global_force:
            raise ValueError(
                "Refusing to delete user templates without explicit purge request."
            )


def _is_template_path(path: str) -> bool:
    normalized = fs_tools.normalize_windows_path(path)
    for template in constants.USER_TEMPLATE_PATHS:
        candidate = fs_tools.normalize_windows_path(template)
        if "%" not in candidate and normalized.startswith(candidate):
            return True
        if candidate.startswith("%APPDATA%\\"):
            suffix = candidate[len("%APPDATA%") :]
            if fs_tools.match_environment_suffix(normalized, "\\APPDATA\\ROAMING" + suffix, require_users=True):
                return True
        if candidate.startswith("%LOCALAPPDATA%\\"):
            suffix = candidate[len("%LOCALAPPDATA%") :]
            if fs_tools.match_environment_suffix(normalized, "\\APPDATA\\LOCAL" + suffix, require_users=True):
                return True
    return False


def _enforce_admin_guard(*, is_admin: bool, dry_run: bool) -> None:
    if dry_run:
        return
    if not is_admin:
        raise PermissionError(
            "Administrative rights are required for destructive operations."
        )


def _enforce_os_guard(
    *,
    os_system: str,
    os_release: str,
    force: bool,
    allow_unsupported_windows: bool,
) -> None:
    system = str(os_system).strip().lower()
    if system and system not in SUPPORTED_SYSTEMS:
        raise RuntimeError(
            f"Unsupported operating system '{os_system}'. Windows is required."
        )

    release_tuple = _parse_windows_release(os_release)
    if (
        release_tuple < MINIMUM_SUPPORTED_WINDOWS_RELEASE
        and not force
        and not allow_unsupported_windows
    ):
        raise RuntimeError(
            "Windows 7 / Server 2008 R2 or newer is required for destructive operations."
        )


def _enforce_process_guard(
    *,
    blocking_processes: Sequence[str],
    dry_run: bool,
    force: bool,
) -> None:
    if dry_run:
        return

    active = [str(proc).strip() for proc in blocking_processes if str(proc).strip()]
    if not active or force:
        return

    joined = ", ".join(active)
    raise RuntimeError(
        "The following Office processes must be closed before continuing: " + joined
    )


def _enforce_restore_point_guard(
    *,
    require_restore_point: bool,
    restore_point_available: bool,
    dry_run: bool,
    force: bool,
) -> None:
    if dry_run or not require_restore_point:
        return

    if restore_point_available or force:
        return

    raise RuntimeError(
        "System restore points are unavailable; re-run with --force or --no-restore-point."
    )


def _parse_windows_release(release: str) -> tuple[int, int]:
    parts: list[int] = []
    for token in str(release).split('.'):
        if len(parts) >= 2:
            break
        token = token.strip()
        if not token:
            continue
        digits = ''.join(ch for ch in token if ch.isdigit())
        if not digits:
            continue
        parts.append(int(digits))
    while len(parts) < 2:
        parts.append(0)
    return parts[0], parts[1]
