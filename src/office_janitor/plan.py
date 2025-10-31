"""!
@brief Translate detection results into actionable scrub plans.
@details Planning resolves requested modes, target Office versions, and
user-selected options into an ordered sequence of steps for uninstall, cleanup,
and backups, matching the workflow outlined in the specification.
"""
from __future__ import annotations

from typing import Dict, Iterable, List, Mapping, MutableMapping, Sequence

from . import constants

_SUPPORTED_TARGETS = tuple(constants.SUPPORTED_TARGETS)
_SUPPORTED_TARGET_SET = {str(value) for value in _SUPPORTED_TARGETS}
_SUPPORTED_COMPONENT_MAP = {item.lower(): item for item in constants.SUPPORTED_COMPONENTS}
_NON_ACTIONABLE_CATEGORIES = {"context", "detect"}

_C2R_RELEASE_HINTS = {
    "o365": "365",
    "365": "365",
    "2024": "2024",
    "2021": "2021",
    "2019": "2019",
    "2016": "2016",
}

_OFFSCRUB_PRIORITY = constants.OFFSCRUB_UNINSTALL_PRIORITY
_MSI_VERSION_GROUPS = constants.MSI_UNINSTALL_VERSION_GROUPS
_C2R_VERSION_GROUPS = constants.C2R_UNINSTALL_VERSION_GROUPS
_DEFAULT_PRIORITY = len(_OFFSCRUB_PRIORITY) + 1

_MSI_MAJOR_VERSION_HINTS = {
    "16": "2016",
    "15": "2013",
    "14": "2010",
    "12": "2007",
    "11": "2003",
}


def build_plan(
    inventory: Dict[str, Sequence[dict]],
    options: Dict[str, object],
    *,
    pass_index: int = 1,
) -> List[dict]:
    """!
    @brief Produce an ordered plan of actions using the current inventory and CLI options.
    @details ``pass_index`` allows the scrubber to regenerate uninstall steps for
    subsequent passes while keeping metadata (such as dependencies) distinct per
    iteration. Cleanup steps remain present in every plan so the executor can run
    them after the final uninstall pass completes.
    """
    normalized_options = _normalize_options(options)
    mode = _resolve_mode(normalized_options)
    dry_run = bool(normalized_options.get("dry_run", False))
    normalized_options["dry_run"] = dry_run
    targets, unsupported_targets = _resolve_targets(mode, normalized_options)
    components, unsupported_components = _resolve_components(normalized_options.get("include"))
    normalized_options["include_components"] = components

    selected_inventory = {
        key: list(value) if not isinstance(value, list) else value
        for key, value in inventory.items()
    }
    discovered_versions = _discover_versions(selected_inventory)
    if not targets:
        targets = discovered_versions

    inventory_summary = _summarize_inventory(selected_inventory, discovered_versions)

    plan: List[dict] = []
    context_metadata = {
        "mode": mode,
        "dry_run": dry_run,
        "force": bool(normalized_options.get("force", False)),
        "target_versions": targets,
        "unsupported_targets": unsupported_targets,
        "discovered_versions": discovered_versions,
        "options": dict(normalized_options),
        "inventory_counts": {
            key: len(value) if hasattr(value, "__len__") else len(list(value))
            for key, value in selected_inventory.items()
        },
        "requested_components": components,
        "unsupported_components": unsupported_components,
        "inventory_summary": inventory_summary,
        "pass_index": int(pass_index),
    }

    plan.append(
        {
            "id": "context",
            "category": "context",
            "description": "Planning context and CLI options.",
            "depends_on": [],
            "metadata": context_metadata,
        }
    )

    detect_step_id = f"detect-{pass_index}-0"
    plan.append(
        {
            "id": detect_step_id,
            "category": "detect",
            "description": "Record detection snapshot for downstream steps.",
            "depends_on": ["context"],
            "metadata": {
                "summary": inventory_summary,
                "dry_run": dry_run,
            },
        }
    )

    diagnose_mode = mode == "diagnose"

    include_uninstalls = (not diagnose_mode) and mode not in {"cleanup-only"}

    uninstall_steps: List[str] = []
    prerequisites = [detect_step_id]

    if include_uninstalls:
        c2r_records = list(
            enumerate(_filter_records_by_target(selected_inventory.get("c2r", []), targets))
        )
        c2r_records.sort(
            key=lambda item: (
                _c2r_uninstall_priority(_infer_version(item[1])),
                item[0],
            )
        )
        for index, (_, record) in enumerate(c2r_records):
            version = _infer_version(record)
            uninstall_id = f"c2r-{pass_index}-{index}"
            plan.append(
                {
                    "id": uninstall_id,
                    "category": "c2r-uninstall",
                    "description": record.get(
                        "description", "Uninstall Click-to-Run packages"
                    ),
                    "depends_on": prerequisites,
                    "metadata": {
                        "installation": record,
                        "version": version,
                        "dry_run": dry_run,
                    },
                }
            )
            uninstall_steps.append(uninstall_id)

        msi_records = list(
            enumerate(_filter_records_by_target(selected_inventory.get("msi", []), targets))
        )
        msi_records.sort(
            key=lambda item: (
                _msi_uninstall_priority(item[1]),
                item[0],
            )
        )
        for index, (_, record) in enumerate(msi_records):
            version = _infer_version(record)
            uninstall_id = f"msi-{pass_index}-{index}"
            plan.append(
                {
                    "id": uninstall_id,
                    "category": "msi-uninstall",
                    "description": record.get(
                        "display_name", f"Uninstall MSI product {record.get('product_code', 'unknown')}"
                    ),
                    "depends_on": prerequisites,
                    "metadata": {
                        "product": record,
                        "version": version,
                        "dry_run": dry_run,
                    },
                }
            )
            uninstall_steps.append(uninstall_id)

    cleanup_dependencies: List[str] = uninstall_steps or [detect_step_id]
    licensing_step_id = ""

    if (not diagnose_mode) and not normalized_options.get("no_license", False):
        licensing_step_id = f"licensing-{pass_index}-0"
        plan.append(
            {
                "id": licensing_step_id,
                "category": "licensing-cleanup",
                "description": "Remove Office licensing and activation tokens.",
                "depends_on": cleanup_dependencies,
                "metadata": {
                    "dry_run": dry_run,
                    "mode": mode,
                },
            }
        )
        cleanup_dependencies = [licensing_step_id]

    task_names = [] if diagnose_mode else _collect_task_names(selected_inventory.get("tasks", []))
    if task_names:
        task_step_id = f"tasks-{pass_index}-0"
        plan.append(
            {
                "id": task_step_id,
                "category": "task-cleanup",
                "description": "Remove Office-related scheduled tasks.",
                "depends_on": cleanup_dependencies,
                "metadata": {
                    "tasks": task_names,
                    "dry_run": dry_run,
                },
            }
        )
        cleanup_dependencies = [task_step_id]

    service_names = [] if diagnose_mode else _collect_service_names(selected_inventory.get("services", []))
    if service_names:
        service_step_id = f"services-{pass_index}-0"
        plan.append(
            {
                "id": service_step_id,
                "category": "service-cleanup",
                "description": "Delete Office background services.",
                "depends_on": cleanup_dependencies,
                "metadata": {
                    "services": service_names,
                    "dry_run": dry_run,
                },
            }
        )
        cleanup_dependencies = [service_step_id]

    filesystem_entries = [] if diagnose_mode else _collect_paths(selected_inventory.get("filesystem", []))
    if filesystem_entries:
        plan.append(
            {
                "id": f"filesystem-{pass_index}-0",
                "category": "filesystem-cleanup",
                "description": "Remove residual Office filesystem artifacts.",
                "depends_on": cleanup_dependencies,
                "metadata": {
                    "paths": filesystem_entries,
                    "preserve_templates": bool(normalized_options.get("keep_templates", False)),
                    "purge_templates": bool(normalized_options.get("force", False))
                    and not bool(normalized_options.get("keep_templates", False)),
                    "dry_run": dry_run,
                },
            }
        )

    registry_entries = [] if diagnose_mode else _collect_registry_paths(selected_inventory.get("registry", []))
    if registry_entries:
        plan.append(
            {
                "id": f"registry-{pass_index}-0",
                "category": "registry-cleanup",
                "description": "Purge Office registry hives and COM registrations.",
                "depends_on": cleanup_dependencies,
                "metadata": {
                    "keys": registry_entries,
                    "dry_run": dry_run,
                },
            }
        )

    summary = summarize_plan(plan)
    context_step = plan[0]
    metadata = dict(context_step.get("metadata", {}))
    metadata["summary"] = summary
    context_step["metadata"] = metadata
    return plan


def _normalize_options(options: Mapping[str, object]) -> Dict[str, object]:
    if hasattr(options, "__dict__"):
        return dict(vars(options))
    return dict(options)


def _resolve_mode(options: Mapping[str, object]) -> str:
    explicit_raw = options.get("mode")
    explicit = str(explicit_raw).strip() if isinstance(explicit_raw, str) else ""
    explicit_lower = explicit.lower()

    if options.get("diagnose") or explicit_lower == "diagnose":
        return "diagnose"
    if options.get("cleanup_only") or explicit_lower == "cleanup-only":
        return "cleanup-only"
    if options.get("auto_all") or explicit_lower == "auto-all":
        return "auto-all"

    target = options.get("target")
    if target:
        return f"target:{target}"

    if explicit_lower.startswith("target:") and len(explicit) > len("target:"):
        return f"target:{explicit.split(":", 1)[1]}"
    if explicit:
        return explicit_lower or explicit

    return "interactive"


def _resolve_targets(mode: str, options: Mapping[str, object]) -> tuple[List[str], List[str]]:
    raw_targets: List[str] = []
    unsupported: List[str] = []

    if mode.startswith("target:"):
        selected = mode.split(":", 1)[1]
        if selected:
            raw_targets.append(str(selected))

    target_option = options.get("target")
    if target_option and str(target_option) not in raw_targets:
        raw_targets.append(str(target_option))

    additional = options.get("targets")
    if additional:
        if isinstance(additional, str):
            raw_targets.extend([item.strip() for item in additional.split(",") if item.strip()])
        elif isinstance(additional, Iterable):
            raw_targets.extend(str(item) for item in additional)

    seen: set[str] = set()
    ordered_targets: List[str] = []
    for candidate in raw_targets:
        candidate_norm = candidate.strip()
        if candidate_norm and candidate_norm not in seen:
            seen.add(candidate_norm)
            ordered_targets.append(candidate_norm)

    for candidate in ordered_targets:
        if candidate not in _SUPPORTED_TARGET_SET:
            unsupported.append(candidate)

    valid_targets = [candidate for candidate in ordered_targets if candidate not in unsupported]
    return valid_targets, unsupported


def _resolve_components(include_option: object) -> tuple[List[str], List[str]]:
    if not include_option:
        return [], []

    raw_components: List[str] = []
    if isinstance(include_option, str):
        raw_components = [item.strip() for item in include_option.split(",") if item.strip()]
    elif isinstance(include_option, Iterable):
        raw_components = [str(item).strip() for item in include_option if str(item).strip()]

    seen: set[str] = set()
    resolved: List[str] = []
    unsupported: List[str] = []
    for candidate in raw_components:
        lower = candidate.lower()
        if lower in seen:
            continue
        seen.add(lower)
        mapped = _SUPPORTED_COMPONENT_MAP.get(lower)
        if mapped:
            resolved.append(mapped)
        else:
            unsupported.append(candidate)
    return resolved, unsupported


def _discover_versions(inventory: Mapping[str, Sequence[dict]]) -> List[str]:
    versions: set[str] = set()
    for key in ("msi", "c2r"):
        for record in inventory.get(key, []):
            version = _infer_version(record)
            if version:
                versions.add(version)
    return sorted(versions)


def _filter_records_by_target(records: Sequence[dict], targets: Sequence[str]) -> List[dict]:
    if not targets:
        return list(records)
    target_set = {str(target) for target in targets}
    filtered: List[dict] = []
    for record in records:
        version = _infer_version(record)
        if version and version in target_set:
            filtered.append(record)
    return filtered


def _infer_version(record: Mapping[str, object]) -> str:
    direct_fields = ("target_version", "version", "major_version", "product_version")
    fallback_value = ""
    for field in direct_fields:
        value = record.get(field)
        if not value:
            continue
        value_str = str(value)
        if value_str in _SUPPORTED_TARGET_SET:
            return value_str
        major_component = value_str.split(".", 1)[0]
        if major_component in _SUPPORTED_TARGET_SET:
            return major_component
        if not fallback_value:
            fallback_value = value_str

    tags = record.get("tags")
    if isinstance(tags, Iterable) and not isinstance(tags, (str, bytes)):
        for tag in tags:
            tag_str = str(tag)
            if tag_str in _SUPPORTED_TARGET_SET:
                return tag_str

    release_ids = record.get("release_ids")
    if isinstance(release_ids, Iterable) and not isinstance(release_ids, (str, bytes)):
        for release in release_ids:
            release_lower = str(release).lower()
            for hint, mapped in _C2R_RELEASE_HINTS.items():
                if hint in release_lower:
                    return mapped

    channel = record.get("channel")
    if isinstance(channel, str):
        for hint, mapped in _C2R_RELEASE_HINTS.items():
            if hint in channel.lower():
                return mapped

    if fallback_value:
        return fallback_value
    return ""


def _collect_paths(entries: Sequence[Mapping[str, object]]) -> List[str]:
    paths: List[str] = []
    for entry in entries:
        candidate = entry.get("path")
        if isinstance(candidate, str) and candidate:
            paths.append(candidate)
    return paths


def _collect_registry_paths(entries: Sequence[Mapping[str, object]]) -> List[str]:
    keys: List[str] = []
    for entry in entries:
        for field in ("key", "path"):
            candidate = entry.get(field)
            if isinstance(candidate, str) and candidate:
                keys.append(candidate)
                break
    return keys


def _collect_task_names(entries: Sequence[Mapping[str, object]]) -> List[str]:
    tasks: List[str] = []
    for entry in entries:
        candidate = entry.get("task") or entry.get("name")
        if isinstance(candidate, str) and candidate:
            tasks.append(candidate)
    return tasks


def _collect_service_names(entries: Sequence[Mapping[str, object]]) -> List[str]:
    services: List[str] = []
    for entry in entries:
        candidate = entry.get("name") or entry.get("service")
        if isinstance(candidate, str) and candidate:
            services.append(candidate)
    return services


def _summarize_inventory(
    inventory: Mapping[str, Sequence[Mapping[str, object]]], discovered_versions: Sequence[str]
) -> Dict[str, object]:
    counts: Dict[str, int] = {}
    total_entries = 0
    for key, items in inventory.items():
        try:
            count = len(items)  # type: ignore[arg-type]
        except TypeError:
            count = len(list(items))
        counts[key] = count
        total_entries += count
    return {
        "counts": counts,
        "total_entries": total_entries,
        "discovered_versions": list(discovered_versions),
    }


def _msi_uninstall_priority(record: Mapping[str, object]) -> int:
    group = _resolve_msi_priority_group(record)
    if not group:
        version = _infer_version(record)
        group = _MSI_VERSION_GROUPS.get(version, version)
    return _OFFSCRUB_PRIORITY.get(group, _DEFAULT_PRIORITY)


def _resolve_msi_priority_group(record: Mapping[str, object]) -> str:
    candidates = _collect_msi_version_candidates(record)
    for candidate in candidates:
        mapped = _MSI_VERSION_GROUPS.get(candidate)
        if mapped:
            return mapped
    for candidate in candidates:
        major = candidate.split(".", 1)[0]
        alias = _MSI_MAJOR_VERSION_HINTS.get(major)
        if not alias:
            continue
        mapped = _MSI_VERSION_GROUPS.get(alias, alias)
        if mapped in _OFFSCRUB_PRIORITY:
            return mapped
    return ""


def _collect_msi_version_candidates(record: Mapping[str, object]) -> List[str]:
    candidates: List[str] = []

    def _add(value: object) -> None:
        if not value:
            return
        text = str(value).strip()
        if not text:
            return
        if text not in candidates:
            candidates.append(text)

    _add(_infer_version(record))
    for field in ("target_version", "version", "major_version", "product_version"):
        _add(record.get(field))

    properties = record.get("properties")
    if isinstance(properties, Mapping):
        for field in ("version", "product_version", "display_version"):
            _add(properties.get(field))
        supported = properties.get("supported_versions")
        if isinstance(supported, Iterable) and not isinstance(supported, (str, bytes)):
            for item in supported:
                _add(item)

    direct_supported = record.get("supported_versions")
    if isinstance(direct_supported, Iterable) and not isinstance(direct_supported, (str, bytes)):
        for item in direct_supported:
            _add(item)

    return candidates


def _c2r_uninstall_priority(version: str) -> int:
    group = _C2R_VERSION_GROUPS.get(version, "c2r")
    return _OFFSCRUB_PRIORITY.get(group, 0)


def summarize_plan(plan_steps: Sequence[Mapping[str, object]]) -> Dict[str, object]:
    """!
    @brief Build a lightweight summary structure for UI and telemetry surfaces.
    @details Aggregates category counts, uninstall targets, and request metadata so
    interactive front-ends can present concise plan details without walking every
    step. Non-actionable categories such as ``context`` and ``detect`` are
    excluded from the actionable step count.
    """

    summary: Dict[str, object] = {
        "total_steps": len(plan_steps),
        "actionable_steps": 0,
        "categories": {},
        "uninstall_versions": [],
        "cleanup_categories": [],
        "mode": "",
        "dry_run": False,
        "target_versions": [],
        "discovered_versions": [],
        "requested_components": [],
        "unsupported_components": [],
        "inventory_counts": {},
    }

    context_metadata: MutableMapping[str, object] | None = None
    categories: Dict[str, int] = {}
    uninstall_versions: set[str] = set()
    cleanup_categories: List[str] = []

    for step in plan_steps:
        category = str(step.get("category", ""))
        categories[category] = categories.get(category, 0) + 1
        if category not in _NON_ACTIONABLE_CATEGORIES:
            summary["actionable_steps"] = int(summary.get("actionable_steps", 0)) + 1
        if category in {"msi-uninstall", "c2r-uninstall"}:
            metadata = step.get("metadata", {})
            if isinstance(metadata, Mapping):
                version = metadata.get("version")
                if version:
                    uninstall_versions.add(str(version))
        if category.endswith("cleanup") and category not in cleanup_categories:
            cleanup_categories.append(category)
        if category == "context" and isinstance(step.get("metadata"), MutableMapping):
            context_metadata = step["metadata"]  # type: ignore[assignment]

    summary["categories"] = categories
    summary["uninstall_versions"] = _sort_versions(uninstall_versions)
    summary["cleanup_categories"] = cleanup_categories

    if context_metadata is not None:
        summary["mode"] = str(context_metadata.get("mode", ""))
        summary["dry_run"] = bool(context_metadata.get("dry_run", False))
        summary["target_versions"] = list(context_metadata.get("target_versions", []))
        summary["discovered_versions"] = list(context_metadata.get("discovered_versions", []))
        summary["requested_components"] = list(context_metadata.get("requested_components", []))
        summary["unsupported_components"] = list(context_metadata.get("unsupported_components", []))
        summary["inventory_counts"] = dict(context_metadata.get("inventory_counts", {}))

    return summary


def _sort_versions(versions: Iterable[str]) -> List[str]:
    order_map = {value: index for index, value in enumerate(_SUPPORTED_TARGETS)}

    def _sort_key(value: str) -> tuple[int, str]:
        lower = value.strip()
        return (order_map.get(lower, len(order_map)), lower)

    return sorted({str(value) for value in versions if value}, key=_sort_key)
