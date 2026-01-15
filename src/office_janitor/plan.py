"""!
@brief Translate detection results into actionable scrub plans.
@details Planning resolves requested modes, target Office versions, and
user-selected options into an ordered sequence of steps for uninstall, cleanup,
and backups, matching the workflow outlined in the specification.
"""

from __future__ import annotations

from collections.abc import Iterable, Mapping, MutableMapping, MutableSequence, Sequence

from .plan_helpers import (
    NON_ACTIONABLE_CATEGORIES,
    augment_auto_all_c2r_inventory,
    c2r_uninstall_priority,
    coerce_to_list,
    coerce_to_mapping,
    collect_paths,
    collect_registry_paths,
    collect_service_names,
    collect_task_names,
    collect_uninstall_handles,
    discover_versions,
    filter_records_by_target,
    infer_version,
    msi_uninstall_priority,
    normalize_options,
    record_matches_product_code_filter,
    record_matches_release_filter,
    resolve_components,
    resolve_mode,
    resolve_targets,
    sort_versions,
    summarize_inventory,
)

# Re-export private names for backward compatibility
_normalize_options = normalize_options
_record_matches_release_filter = record_matches_release_filter
_record_matches_product_code_filter = record_matches_product_code_filter
_resolve_mode = resolve_mode
_resolve_targets = resolve_targets
_resolve_components = resolve_components
_augment_auto_all_c2r_inventory = augment_auto_all_c2r_inventory
_discover_versions = discover_versions
_filter_records_by_target = filter_records_by_target
_infer_version = infer_version
_collect_paths = collect_paths
_collect_registry_paths = collect_registry_paths
_collect_uninstall_handles = collect_uninstall_handles
_collect_task_names = collect_task_names
_collect_service_names = collect_service_names
_summarize_inventory = summarize_inventory
_msi_uninstall_priority = msi_uninstall_priority
_c2r_uninstall_priority = c2r_uninstall_priority
_sort_versions = sort_versions
_coerce_to_list = coerce_to_list
_coerce_to_mapping = coerce_to_mapping


def build_plan(
    inventory: Mapping[str, Sequence[Mapping[str, object]]],
    options: Mapping[str, object],
    *,
    pass_index: int = 1,
) -> list[dict[str, object]]:
    """!
    @brief Produce an ordered plan of actions using the current inventory and CLI options.
    @details ``pass_index`` allows the scrubber to regenerate uninstall steps for
    subsequent passes while keeping metadata (such as dependencies) distinct per
    iteration. Cleanup steps remain present in every plan so the executor can run
    them after the final uninstall pass completes.
    """
    normalized_options = normalize_options(options)
    mode = resolve_mode(normalized_options)
    dry_run = bool(normalized_options.get("dry_run", False))
    normalized_options["dry_run"] = dry_run
    targets, unsupported_targets = resolve_targets(mode, normalized_options)
    components, unsupported_components = resolve_components(normalized_options.get("include"))
    normalized_options["include_components"] = components

    detected_inventory: dict[str, list[Mapping[str, object]]] = {}
    for key, value in inventory.items():
        records: list[Mapping[str, object]] = []
        if isinstance(value, list):
            records = list(value)
        elif isinstance(value, Iterable) and not isinstance(value, (str, bytes)):
            records = list(value)
        detected_inventory[key] = records

    planning_inventory: MutableMapping[str, MutableSequence[Mapping[str, object]]] = {
        key: list(value) for key, value in detected_inventory.items()
    }
    if mode == "auto-all":
        augment_auto_all_c2r_inventory(planning_inventory, components)

    detected_versions = discover_versions(detected_inventory)
    planning_versions = discover_versions(planning_inventory)
    if not targets:
        targets = planning_versions

    inventory_summary = summarize_inventory(detected_inventory, detected_versions)

    plan: list[dict[str, object]] = []
    context_metadata = {
        "mode": mode,
        "dry_run": dry_run,
        "force": bool(normalized_options.get("force", False)),
        "target_versions": targets,
        "unsupported_targets": unsupported_targets,
        "discovered_versions": detected_versions,
        "options": dict(normalized_options),
        "inventory_counts": {
            key: len(value) if hasattr(value, "__len__") else len(list(value))
            for key, value in detected_inventory.items()
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

    # Extract uninstall method filtering options
    uninstall_method = str(normalized_options.get("uninstall_method", "auto")).lower()
    include_c2r = uninstall_method in ("auto", "c2r", "click-to-run")
    include_msi = uninstall_method in ("auto", "msi")
    product_code_filter = normalized_options.get("product_codes")
    release_id_filter = normalized_options.get("release_ids")

    # Extract retry options for step metadata
    retries = int(normalized_options.get("retries", 9) or 9)
    retry_delay = int(normalized_options.get("retry_delay", 3) or 3)
    retry_delay_max = int(normalized_options.get("retry_delay_max", 30) or 30)
    force_app_shutdown = bool(normalized_options.get("force_app_shutdown", False))

    uninstall_steps: list[str] = []
    prerequisites = [detect_step_id]

    if include_uninstalls:
        # C2R uninstall steps (if not filtered out)
        if include_c2r:
            c2r_records = list(
                enumerate(filter_records_by_target(planning_inventory.get("c2r", []), targets))
            )
            # Apply release_id filter if specified
            if release_id_filter:
                filter_set = set(
                    str(rid).strip().lower()
                    for rid in (
                        release_id_filter
                        if isinstance(release_id_filter, list)
                        else [release_id_filter]
                    )
                )
                c2r_records = [
                    (idx, rec)
                    for idx, rec in c2r_records
                    if record_matches_release_filter(rec, filter_set)
                ]
            c2r_records.sort(
                key=lambda item: (
                    c2r_uninstall_priority(infer_version(item[1])),
                    item[0],
                )
            )
            for index, (_, record) in enumerate(c2r_records):
                version = infer_version(record)
                uninstall_id = f"c2r-{pass_index}-{index}"
                plan.append(
                    {
                        "id": uninstall_id,
                        "category": "c2r-uninstall",
                        "description": record.get("description", "Uninstall Click-to-Run packages"),
                        "depends_on": prerequisites,
                        "metadata": {
                            "installation": record,
                            "version": version,
                            "dry_run": dry_run,
                            "force": force_app_shutdown,
                            "retries": retries,
                            "retry_delay": retry_delay,
                            "retry_delay_max": retry_delay_max,
                        },
                    }
                )
                uninstall_steps.append(uninstall_id)

        # MSI uninstall steps (if not filtered out)
        if include_msi:
            msi_records = list(
                enumerate(filter_records_by_target(planning_inventory.get("msi", []), targets))
            )
            # Apply product_code filter if specified
            if product_code_filter:
                filter_set = set(
                    str(pc).strip().upper()
                    for pc in (
                        product_code_filter
                        if isinstance(product_code_filter, list)
                        else [product_code_filter]
                    )
                )
                msi_records = [
                    (idx, rec)
                    for idx, rec in msi_records
                    if record_matches_product_code_filter(rec, filter_set)
                ]
            msi_records.sort(
                key=lambda item: (
                    msi_uninstall_priority(item[1]),
                    item[0],
                )
            )
            for index, (_, record) in enumerate(msi_records):
                version = infer_version(record)
                uninstall_id = f"msi-{pass_index}-{index}"
                plan.append(
                    {
                        "id": uninstall_id,
                        "category": "msi-uninstall",
                        "description": record.get(
                            "display_name",
                            f"Uninstall MSI product {record.get('product_code', 'unknown')}",
                        ),
                        "depends_on": prerequisites,
                        "metadata": {
                            "product": record,
                            "version": version,
                            "dry_run": dry_run,
                            "force": force_app_shutdown,
                            "retries": retries,
                            "retry_delay": retry_delay,
                            "retry_delay_max": retry_delay_max,
                        },
                    }
                )
                uninstall_steps.append(uninstall_id)

    cleanup_dependencies: list[str] = uninstall_steps or [detect_step_id]
    licensing_step_id = ""

    # Extract skip flags
    skip_tasks = bool(normalized_options.get("skip_tasks", False))
    skip_services = bool(normalized_options.get("skip_services", False))
    skip_filesystem = bool(normalized_options.get("skip_filesystem", False))
    skip_registry = bool(normalized_options.get("skip_registry", False))

    # Scrub level determines default cleanup intensity
    scrub_level = str(normalized_options.get("scrub_level", "standard")).lower()
    is_aggressive = scrub_level in ("aggressive", "nuclear")
    is_nuclear = scrub_level == "nuclear"

    # For minimal scrub, skip all cleanup
    if scrub_level == "minimal":
        skip_tasks = True
        skip_services = True
        skip_filesystem = True
        skip_registry = True

    # Extract license cleanup granularity flags
    clean_spp = is_nuclear or bool(normalized_options.get("clean_spp", False))
    clean_ospp = is_nuclear or bool(normalized_options.get("clean_ospp", False))
    clean_vnext = is_nuclear or bool(normalized_options.get("clean_vnext", False))
    clean_all_licenses = is_nuclear or bool(normalized_options.get("clean_all_licenses", False))

    # Extract extended cleanup flags
    clean_msocache = is_nuclear or bool(normalized_options.get("clean_msocache", False))
    clean_appx = is_nuclear or bool(normalized_options.get("clean_appx", False))
    clean_wi_metadata = is_nuclear or bool(normalized_options.get("clean_wi_metadata", False))
    clean_shortcuts = is_aggressive or bool(normalized_options.get("clean_shortcuts", False))

    # Extract registry cleanup granularity flags
    clean_addin_registry = is_aggressive or bool(
        normalized_options.get("clean_addin_registry", False)
    )
    clean_com_registry = is_aggressive or bool(normalized_options.get("clean_com_registry", False))
    clean_shell_extensions = is_aggressive or bool(
        normalized_options.get("clean_shell_extensions", False)
    )
    clean_typelibs = is_nuclear or bool(normalized_options.get("clean_typelibs", False))
    clean_protocol_handlers = is_nuclear or bool(
        normalized_options.get("clean_protocol_handlers", False)
    )
    remove_vba = is_nuclear or bool(normalized_options.get("remove_vba", False))

    if (not diagnose_mode) and not (
        normalized_options.get("no_license", False) or normalized_options.get("keep_license", False)
    ):
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
                    "clean_spp": clean_spp,
                    "clean_ospp": clean_ospp,
                    "clean_vnext": clean_vnext,
                    "clean_all_licenses": clean_all_licenses,
                    "retries": retries,
                    "retry_delay": retry_delay,
                },
            }
        )
        cleanup_dependencies = [licensing_step_id]

    # Task cleanup (unless skipped)
    task_names = (
        []
        if (diagnose_mode or skip_tasks)
        else collect_task_names(planning_inventory.get("tasks", []))
    )
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
                    "retries": retries,
                    "retry_delay": retry_delay,
                },
            }
        )
        cleanup_dependencies = [task_step_id]

    # Service cleanup (unless skipped)
    service_names = (
        []
        if (diagnose_mode or skip_services)
        else collect_service_names(planning_inventory.get("services", []))
    )
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
                    "retries": retries,
                    "retry_delay": retry_delay,
                },
            }
        )
        cleanup_dependencies = [service_step_id]

    # Filesystem cleanup (unless skipped)
    filesystem_entries = (
        []
        if (diagnose_mode or skip_filesystem)
        else collect_paths(planning_inventory.get("filesystem", []))
    )
    if filesystem_entries or clean_msocache or clean_appx or clean_shortcuts:
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
                    "clean_msocache": clean_msocache,
                    "clean_appx": clean_appx,
                    "clean_shortcuts": clean_shortcuts,
                    "retries": retries,
                    "retry_delay": retry_delay,
                },
            }
        )
        cleanup_dependencies = [f"filesystem-{pass_index}-0"]

    # Registry cleanup (unless skipped)
    registry_entries = (
        []
        if (diagnose_mode or skip_registry)
        else collect_registry_paths(planning_inventory.get("registry", []))
    )

    # Include uninstall registry entries in cleanup
    if not diagnose_mode and not skip_registry:
        uninstall_handles = collect_uninstall_handles(
            planning_inventory.get("uninstall_entries", [])
        )
        for handle in uninstall_handles:
            if handle not in registry_entries:
                registry_entries.append(handle)

    has_extended_registry = any(
        [
            clean_addin_registry,
            clean_com_registry,
            clean_shell_extensions,
            clean_typelibs,
            clean_protocol_handlers,
            remove_vba,
            clean_wi_metadata,
        ]
    )
    if registry_entries or has_extended_registry:
        plan.append(
            {
                "id": f"registry-{pass_index}-0",
                "category": "registry-cleanup",
                "description": "Purge Office registry hives and COM registrations.",
                "depends_on": cleanup_dependencies,
                "metadata": {
                    "keys": registry_entries,
                    "dry_run": dry_run,
                    "clean_addin_registry": clean_addin_registry,
                    "clean_com_registry": clean_com_registry,
                    "clean_shell_extensions": clean_shell_extensions,
                    "clean_typelibs": clean_typelibs,
                    "clean_protocol_handlers": clean_protocol_handlers,
                    "remove_vba": remove_vba,
                    "clean_wi_metadata": clean_wi_metadata,
                    "retries": retries,
                    "retry_delay": retry_delay,
                },
            }
        )
        cleanup_dependencies = [f"registry-{pass_index}-0"]

    # vNext identity cleanup (aggressive/nuclear or explicit)
    if not diagnose_mode and (is_aggressive or clean_vnext):
        plan.append(
            {
                "id": f"vnext-identity-{pass_index}-0",
                "category": "vnext-identity-cleanup",
                "description": "Clean vNext identity and device licensing registry.",
                "depends_on": cleanup_dependencies,
                "metadata": {
                    "dry_run": dry_run,
                    "retries": retries,
                    "retry_delay": retry_delay,
                },
            }
        )
        cleanup_dependencies = [f"vnext-identity-{pass_index}-0"]

    # Taskband cleanup (nuclear or explicit)
    if not diagnose_mode and is_nuclear:
        plan.append(
            {
                "id": f"taskband-{pass_index}-0",
                "category": "taskband-cleanup",
                "description": "Clean Office pinned items from taskbar.",
                "depends_on": cleanup_dependencies,
                "metadata": {
                    "include_all_users": True,
                    "dry_run": dry_run,
                    "retries": retries,
                    "retry_delay": retry_delay,
                },
            }
        )
        cleanup_dependencies = [f"taskband-{pass_index}-0"]

    # Published components cleanup (nuclear only)
    if not diagnose_mode and is_nuclear:
        plan.append(
            {
                "id": f"published-components-{pass_index}-0",
                "category": "published-components-cleanup",
                "description": "Clean Office entries from Windows Installer published components.",
                "depends_on": cleanup_dependencies,
                "metadata": {
                    "dry_run": dry_run,
                    "retries": retries,
                    "retry_delay": retry_delay,
                },
            }
        )

    summary = summarize_plan(plan)
    context_step = plan[0]
    metadata = dict(context_step.get("metadata", {}))
    metadata["summary"] = summary
    context_step["metadata"] = metadata
    return plan


def summarize_plan(plan_steps: Sequence[Mapping[str, object]]) -> dict[str, object]:
    """!
    @brief Build a lightweight summary structure for UI and telemetry surfaces.
    @details Aggregates category counts, uninstall targets, and request metadata so
    interactive front-ends can present concise plan details without walking every
    step. Non-actionable categories such as ``context`` and ``detect`` are
    excluded from the actionable step count.
    """

    summary: dict[str, object] = {
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
    categories: dict[str, int] = {}
    uninstall_versions: set[str] = set()
    cleanup_categories: list[str] = []

    for step in plan_steps:
        category = str(step.get("category", ""))
        categories[category] = categories.get(category, 0) + 1
        if category not in NON_ACTIONABLE_CATEGORIES:
            actionable_value = summary.get("actionable_steps")
            actionable_steps = actionable_value if isinstance(actionable_value, int) else 0
            summary["actionable_steps"] = actionable_steps + 1
        if category in {"msi-uninstall", "c2r-uninstall"}:
            metadata = step.get("metadata", {})
            if isinstance(metadata, Mapping):
                version = metadata.get("version")
                if version:
                    uninstall_versions.add(str(version))
        if category.endswith("cleanup") and category not in cleanup_categories:
            cleanup_categories.append(category)
        metadata_obj = step.get("metadata")
        if category == "context" and isinstance(metadata_obj, MutableMapping):
            context_metadata = metadata_obj

    summary["categories"] = categories
    summary["uninstall_versions"] = sort_versions(uninstall_versions)
    summary["cleanup_categories"] = cleanup_categories

    if context_metadata is not None:
        summary["mode"] = str(context_metadata.get("mode", ""))
        summary["dry_run"] = bool(context_metadata.get("dry_run", False))
        summary["target_versions"] = coerce_to_list(context_metadata.get("target_versions"))
        summary["discovered_versions"] = coerce_to_list(
            context_metadata.get("discovered_versions")
        )
        summary["requested_components"] = coerce_to_list(
            context_metadata.get("requested_components")
        )
        summary["unsupported_components"] = coerce_to_list(
            context_metadata.get("unsupported_components")
        )
        summary["inventory_counts"] = coerce_to_mapping(context_metadata.get("inventory_counts"))

    return summary
