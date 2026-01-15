"""!
@file plan_helpers.py
@brief Helper functions for plan building and resolution.

@details Contains private helper functions for option normalization,
target/component resolution, inventory augmentation, version inference,
record filtering, and priority sorting.
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping, MutableMapping, MutableSequence, Sequence

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


def normalize_options(options: Mapping[str, object]) -> dict[str, object]:
    """!
    @brief Convert options mapping to a mutable dictionary.
    """
    if hasattr(options, "__dict__"):
        return dict(vars(options))
    return dict(options)


def record_matches_release_filter(record: Mapping[str, object], filter_set: set[str]) -> bool:
    """!
    @brief Check if a C2R record matches any release_id in the filter set.
    """
    release_ids = record.get("release_ids")
    if isinstance(release_ids, Iterable) and not isinstance(release_ids, (str, bytes)):
        for rid in release_ids:
            if str(rid).strip().lower() in filter_set:
                return True
    single_id = record.get("release_id")
    if single_id and str(single_id).strip().lower() in filter_set:
        return True
    return False


def record_matches_product_code_filter(record: Mapping[str, object], filter_set: set[str]) -> bool:
    """!
    @brief Check if an MSI record matches any product_code in the filter set.
    """
    product_code = record.get("product_code")
    if product_code:
        normalized = str(product_code).strip().upper()
        # Handle both with and without braces
        if normalized in filter_set:
            return True
        if normalized.startswith("{") and normalized.endswith("}"):
            if normalized[1:-1] in filter_set:
                return True
        else:
            if f"{{{normalized}}}" in filter_set:
                return True
    return False


def resolve_mode(options: Mapping[str, object]) -> str:
    """!
    @brief Determine the planning mode from options.
    """
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
        return f"target:{explicit.split(':', 1)[1]}"
    if explicit:
        return explicit_lower or explicit

    return "interactive"


def resolve_targets(mode: str, options: Mapping[str, object]) -> tuple[list[str], list[str]]:
    """!
    @brief Extract and validate target versions from mode and options.
    @return Tuple of (valid_targets, unsupported_targets).
    """
    raw_targets: list[str] = []
    unsupported: list[str] = []

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
    ordered_targets: list[str] = []
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


def resolve_components(include_option: object) -> tuple[list[str], list[str]]:
    """!
    @brief Parse and validate component include option.
    @return Tuple of (resolved_components, unsupported_components).
    """
    if not include_option:
        return [], []

    raw_components: list[str] = []
    if isinstance(include_option, str):
        raw_components = [item.strip() for item in include_option.split(",") if item.strip()]
    elif isinstance(include_option, Iterable):
        raw_components = [str(item).strip() for item in include_option if str(item).strip()]

    seen: set[str] = set()
    resolved: list[str] = []
    unsupported: list[str] = []
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


def augment_auto_all_c2r_inventory(
    inventory: MutableMapping[str, MutableSequence[Mapping[str, object]]],
    include_components: Sequence[str],
) -> None:
    """!
    @brief Add seeded C2R entries for auto-all mode.
    """
    bucket = inventory.get("c2r")
    if bucket is None:
        records: list[Mapping[str, object]] = []
    elif isinstance(bucket, list):
        records = bucket
    else:
        records = list(bucket)
    inventory["c2r"] = records

    include_set = {component.lower() for component in include_components}
    optional_families = {"project", "visio", "onenote"}

    existing_ids: set[str] = set()
    for record in records:
        release_ids = record.get("release_ids")
        if isinstance(release_ids, Iterable) and not isinstance(release_ids, (str, bytes)):
            for release_id in release_ids:
                release_text = str(release_id).strip().lower()
                if release_text:
                    existing_ids.add(release_text)

    for release_id, metadata in constants.DEFAULT_AUTO_ALL_C2R_RELEASES.items():
        base_metadata = constants.C2R_PRODUCT_RELEASES.get(release_id, {})
        family = str(metadata.get("family") or base_metadata.get("family") or "office").lower()
        if family in optional_families and family not in include_set:
            continue
        canonical = release_id.lower()
        if canonical in existing_ids:
            continue
        seeded_entry = _build_seeded_c2r_entry(release_id, metadata, base_metadata)
        records.append(seeded_entry)
        existing_ids.add(canonical)


def _build_seeded_c2r_entry(
    release_id: str,
    metadata: Mapping[str, object],
    base_metadata: Mapping[str, object],
) -> dict[str, object]:
    """!
    @brief Construct a seeded C2R inventory entry from metadata.
    """
    product_name = str(metadata.get("product") or base_metadata.get("product") or release_id)
    description = str(metadata.get("description") or f"Uninstall {product_name}")

    supported_versions_source = metadata.get("supported_versions") or base_metadata.get(
        "supported_versions", ()
    )
    if isinstance(supported_versions_source, (str, bytes)):
        supported_versions_iter: Iterable[object] = [supported_versions_source]
    elif isinstance(supported_versions_source, Iterable):
        supported_versions_iter = supported_versions_source
    else:
        supported_versions_iter = []

    supported_versions: list[str] = []
    for item in supported_versions_iter:
        text = str(item).strip()
        if text:
            supported_versions.append(text)

    default_version = str(metadata.get("default_version") or "").strip()
    if not default_version:
        fallback_version = metadata.get("version")
        if fallback_version:
            default_version = str(fallback_version).strip()
    if not default_version and supported_versions:
        default_version = str(supported_versions[-1]).strip()
    if not default_version:
        default_version = "365" if release_id.lower().startswith("o365") else "c2r"

    if default_version and default_version not in supported_versions:
        supported_versions.append(default_version)

    architectures_source = metadata.get("architectures") or base_metadata.get("architectures", ())
    if isinstance(architectures_source, (str, bytes)):
        architectures_iter: Iterable[object] = [architectures_source]
    elif isinstance(architectures_source, Iterable):
        architectures_iter = architectures_source
    else:
        architectures_iter = []

    supported_architectures: list[str] = []
    for item in architectures_iter:
        text = str(item).strip()
        if text:
            supported_architectures.append(text)
    architecture = str(
        metadata.get("architecture")
        or (supported_architectures[0] if supported_architectures else "x64")
    )

    family = str(metadata.get("family") or base_metadata.get("family") or "").strip()
    channel = str(metadata.get("channel") or base_metadata.get("channel") or "unknown")

    properties: dict[str, object] = {
        "release_id": release_id,
        "version": default_version,
        "supported_versions": supported_versions,
    }
    if supported_architectures:
        properties["supported_architectures"] = supported_architectures
    if family:
        properties["family"] = family
    if channel and channel != "unknown":
        properties["channel"] = channel

    tags: list[str] = []
    raw_tags = metadata.get("tags")
    if isinstance(raw_tags, Iterable) and not isinstance(raw_tags, (str, bytes)):
        for tag in raw_tags:
            tag_text = str(tag).strip()
            if tag_text and tag_text not in tags:
                tags.append(tag_text)
    if default_version and default_version not in tags:
        tags.append(default_version)

    seeded: dict[str, object] = {
        "source": "C2R",
        "product": product_name,
        "description": description,
        "version": default_version,
        "architecture": architecture,
        "release_ids": [release_id],
        "channel": channel,
        "uninstall_handles": [],
        "properties": properties,
    }
    if tags:
        seeded["tags"] = tags

    return seeded


def discover_versions(inventory: Mapping[str, Sequence[Mapping[str, object]]]) -> list[str]:
    """!
    @brief Extract all Office versions from inventory records.
    """
    versions: set[str] = set()
    for key in ("msi", "c2r"):
        for record in inventory.get(key, []):
            version = infer_version(record)
            if version:
                versions.add(version)
    return sorted(versions)


def filter_records_by_target(
    records: Sequence[Mapping[str, object]], targets: Sequence[str]
) -> list[Mapping[str, object]]:
    """!
    @brief Filter records to those matching target versions.
    """
    if not targets:
        return list(records)
    target_set = {str(target) for target in targets}
    filtered: list[Mapping[str, object]] = []
    for record in records:
        version = infer_version(record)
        if version and version in target_set:
            filtered.append(record)
    return filtered


def infer_version(record: Mapping[str, object]) -> str:
    """!
    @brief Derive the Office version from a detection record.
    """
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


def collect_paths(entries: Sequence[Mapping[str, object]]) -> list[str]:
    """!
    @brief Extract filesystem paths from inventory entries.
    """
    paths: list[str] = []
    for entry in entries:
        candidate = entry.get("path")
        if isinstance(candidate, str) and candidate:
            paths.append(candidate)
    return paths


def collect_registry_paths(entries: Sequence[Mapping[str, object]]) -> list[str]:
    """!
    @brief Extract registry key paths from inventory entries.
    """
    keys: list[str] = []
    for entry in entries:
        for field in ("key", "path"):
            candidate = entry.get(field)
            if isinstance(candidate, str) and candidate:
                keys.append(candidate)
                break
    return keys


def collect_uninstall_handles(entries: Sequence[Mapping[str, object]]) -> list[str]:
    """!
    @brief Extract registry handles from detected uninstall entries.
    @details Each uninstall entry from detect_uninstall_entries() has a
    ``registry_handle`` field like ``HKLM\\SOFTWARE\\Microsoft\\Windows\\
    CurrentVersion\\Uninstall\\{product-guid}``. These are the Control Panel
    entries that the VBS scrubber explicitly removes.
    """
    handles: list[str] = []
    for entry in entries:
        candidate = entry.get("registry_handle")
        if isinstance(candidate, str) and candidate:
            handles.append(candidate)
    return handles


def collect_task_names(entries: Sequence[Mapping[str, object]]) -> list[str]:
    """!
    @brief Extract scheduled task names from inventory entries.
    """
    tasks: list[str] = []
    for entry in entries:
        candidate = entry.get("task") or entry.get("name")
        if isinstance(candidate, str) and candidate:
            tasks.append(candidate)
    return tasks


def collect_service_names(entries: Sequence[Mapping[str, object]]) -> list[str]:
    """!
    @brief Extract service names from inventory entries.
    """
    services: list[str] = []
    for entry in entries:
        candidate = entry.get("name") or entry.get("service")
        if isinstance(candidate, str) and candidate:
            services.append(candidate)
    return services


def summarize_inventory(
    inventory: Mapping[str, Sequence[Mapping[str, object]]], discovered_versions: Sequence[str]
) -> dict[str, object]:
    """!
    @brief Build inventory summary for plan metadata.
    """
    counts: dict[str, int] = {}
    total_entries = 0
    for key, items in inventory.items():
        count = len(items)
        counts[key] = count
        total_entries += count
    return {
        "counts": counts,
        "total_entries": total_entries,
        "discovered_versions": list(discovered_versions),
    }


def msi_uninstall_priority(record: Mapping[str, object]) -> int:
    """!
    @brief Compute uninstall priority for an MSI record.
    """
    group = _resolve_msi_priority_group(record)
    if not group:
        version = infer_version(record)
        group = _MSI_VERSION_GROUPS.get(version, version)
    return _OFFSCRUB_PRIORITY.get(group, _DEFAULT_PRIORITY)


def _resolve_msi_priority_group(record: Mapping[str, object]) -> str:
    """!
    @brief Determine the priority group for an MSI record.
    """
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


def _collect_msi_version_candidates(record: Mapping[str, object]) -> list[str]:
    """!
    @brief Gather all possible version strings from an MSI record.
    """
    candidates: list[str] = []

    def _add(value: object) -> None:
        if not value:
            return
        text = str(value).strip()
        if not text:
            return
        if text not in candidates:
            candidates.append(text)

    _add(infer_version(record))
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


def c2r_uninstall_priority(version: str) -> int:
    """!
    @brief Compute uninstall priority for a C2R version.
    """
    group = _C2R_VERSION_GROUPS.get(version, "c2r")
    return _OFFSCRUB_PRIORITY.get(group, 0)


def sort_versions(versions: Iterable[str]) -> list[str]:
    """!
    @brief Sort version strings in canonical order.
    """
    order_map = {value: index for index, value in enumerate(_SUPPORTED_TARGETS)}

    def _sort_key(value: str) -> tuple[int, str]:
        lower = value.strip()
        return (order_map.get(lower, len(order_map)), lower)

    return sorted({str(value) for value in versions if value}, key=_sort_key)


def coerce_to_list(value: object) -> list[str]:
    """!
    @brief Convert value to a list of strings.
    """
    if isinstance(value, list):
        return [str(item) for item in value]
    if isinstance(value, Iterable) and not isinstance(value, (str, bytes)):
        return [str(item) for item in value]
    return []


def coerce_to_mapping(value: object) -> dict[str, object]:
    """!
    @brief Convert value to a dictionary.
    """
    if isinstance(value, dict):
        return dict(value)
    if isinstance(value, Mapping):
        return dict(value)
    return {}


# Re-export constants for use by plan.py
SUPPORTED_TARGET_SET = _SUPPORTED_TARGET_SET
NON_ACTIONABLE_CATEGORIES = _NON_ACTIONABLE_CATEGORIES
