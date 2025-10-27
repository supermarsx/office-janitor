"""!
@brief Safety and guardrail enforcement.
@details Implements dry-run, whitelist/blacklist checks, and preflight
validation to keep cleanup actions safe as called for in the specification.
"""
from __future__ import annotations

from typing import Iterable, Mapping, Sequence

from . import constants, fs_tools

FILESYSTEM_WHITELIST = fs_tools.FILESYSTEM_WHITELIST

FILESYSTEM_BLACKLIST = fs_tools.FILESYSTEM_BLACKLIST

REGISTRY_WHITELIST = (
    r"HKLM\\SOFTWARE\\MICROSOFT\\OFFICE",
    r"HKLM\\SOFTWARE\\WOW6432NODE\\MICROSOFT\\OFFICE",
    r"HKLM\\SOFTWARE\\MICROSOFT\\CLICKTORUN",
    r"HKCU\\SOFTWARE\\MICROSOFT\\OFFICE",
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
