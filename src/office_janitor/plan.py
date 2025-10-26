"""!
@brief Translate detection results into actionable scrub plans.
@details Planning resolves requested modes, target Office versions, and
user-selected options into an ordered sequence of steps for uninstall, cleanup,
and backups, matching the workflow outlined in the specification.
"""
from __future__ import annotations

from typing import Dict, Iterable, List, Mapping, Sequence

SUPPORTED_TARGET_VERSIONS = {
    "2003",
    "2007",
    "2010",
    "2013",
    "2016",
    "2019",
    "2021",
    "2024",
    "365",
}

_C2R_RELEASE_HINTS = {
    "o365": "365",
    "365": "365",
    "2024": "2024",
    "2021": "2021",
    "2019": "2019",
    "2016": "2016",
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
    targets, unsupported = _resolve_targets(mode, normalized_options)

    selected_inventory = {
        key: list(value) if not isinstance(value, list) else value
        for key, value in inventory.items()
    }
    discovered_versions = _discover_versions(selected_inventory)
    if not targets:
        targets = discovered_versions

    plan: List[dict] = []
    context_metadata = {
        "mode": mode,
        "dry_run": dry_run,
        "force": bool(normalized_options.get("force", False)),
        "target_versions": targets,
        "unsupported_targets": unsupported,
        "discovered_versions": discovered_versions,
        "options": dict(normalized_options),
        "inventory_counts": {
            key: len(value) if hasattr(value, "__len__") else len(list(value))
            for key, value in selected_inventory.items()
        },
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

    if mode == "diagnose":
        return plan

    include_uninstalls = mode not in {"cleanup-only"}

    uninstall_steps: List[str] = []

    if include_uninstalls:
        for index, record in enumerate(
            _filter_records_by_target(selected_inventory.get("msi", []), targets)
        ):
            version = _infer_version(record)
            uninstall_id = f"msi-{pass_index}-{index}"
            plan.append(
                {
                    "id": uninstall_id,
                    "category": "msi-uninstall",
                    "description": record.get(
                        "display_name", f"Uninstall MSI product {record.get('product_code', 'unknown')}"
                    ),
                    "depends_on": ["context"],
                    "metadata": {
                        "product": record,
                        "version": version,
                        "dry_run": dry_run,
                    },
                }
            )
            uninstall_steps.append(uninstall_id)

        c2r_records = _filter_records_by_target(selected_inventory.get("c2r", []), targets)
        for index, record in enumerate(c2r_records):
            version = _infer_version(record)
            uninstall_id = f"c2r-{pass_index}-{index}"
            plan.append(
                {
                    "id": uninstall_id,
                    "category": "c2r-uninstall",
                    "description": record.get(
                        "description", "Uninstall Click-to-Run packages"
                    ),
                    "depends_on": ["context"],
                    "metadata": {
                        "installation": record,
                        "version": version,
                        "dry_run": dry_run,
                    },
                }
            )
            uninstall_steps.append(uninstall_id)

    cleanup_dependencies: List[str] = uninstall_steps or ["context"]
    licensing_step_id = ""

    if not normalized_options.get("no_license", False):
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

    filesystem_entries = _collect_paths(selected_inventory.get("filesystem", []))
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

    registry_entries = _collect_registry_paths(selected_inventory.get("registry", []))
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
        if candidate not in SUPPORTED_TARGET_VERSIONS:
            unsupported.append(candidate)

    valid_targets = [candidate for candidate in ordered_targets if candidate not in unsupported]
    return valid_targets, unsupported


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
        if value_str in SUPPORTED_TARGET_VERSIONS:
            return value_str
        major_component = value_str.split(".", 1)[0]
        if major_component in SUPPORTED_TARGET_VERSIONS:
            return major_component
        if not fallback_value:
            fallback_value = value_str

    tags = record.get("tags")
    if isinstance(tags, Iterable) and not isinstance(tags, (str, bytes)):
        for tag in tags:
            tag_str = str(tag)
            if tag_str in SUPPORTED_TARGET_VERSIONS:
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
