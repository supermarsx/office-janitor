"""!
@brief Orchestrate uninstallation, cleanup, and reporting steps.
@details The scrubber consumes an action plan and coordinates MSI/C2R uninstall
routines, license cleanup, filesystem and registry purges, and telemetry
emission as laid out in the specification.
"""
from __future__ import annotations

from typing import Iterable, Mapping

from . import (
    c2r_uninstall,
    constants,
    fs_tools,
    licensing,
    logging_ext,
    msi_uninstall,
    processes,
    restore_point,
    tasks_services,
)


def execute_plan(plan: Iterable[Mapping[str, object]], *, dry_run: bool = False) -> None:
    """!
    @brief Run each plan step while respecting dry-run safety requirements.
    """
    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    steps = [dict(step) for step in plan]
    if not steps:
        human_logger.info("No plan steps supplied; nothing to execute.")
        return

    context = next((step for step in steps if step.get("category") == "context"), None)
    context_metadata = dict(context.get("metadata", {})) if context else {}
    options = dict(context_metadata.get("options", {})) if context_metadata else {}

    global_dry_run = bool(dry_run or context_metadata.get("dry_run", False))

    machine_logger.info(
        "scrub_plan_start",
        extra={
            "event": "scrub_plan_start",
            "step_count": len(steps),
            "dry_run": global_dry_run,
            "options": options,
        },
    )

    if global_dry_run:
        human_logger.info("Executing plan in dry-run mode; no destructive actions will occur.")
    else:
        if options.get("create_restore_point") or options.get("restore_point"):
            try:
                restore_point.create_restore_point("Office Janitor pre-cleanup")
            except Exception as exc:  # pragma: no cover - defensive logging
                human_logger.warning("Failed to create restore point: %s", exc)

        processes.terminate_office_processes(constants.DEFAULT_OFFICE_PROCESSES)
        tasks_services.stop_services(constants.KNOWN_SERVICES)
        tasks_services.disable_tasks(constants.KNOWN_SCHEDULED_TASKS, dry_run=False)

    for step in steps:
        category = step.get("category", "unknown")
        metadata = dict(step.get("metadata", {}))
        step_dry_run = global_dry_run or bool(metadata.get("dry_run", False))

        machine_logger.info(
            "scrub_step_start",
            extra={
                "event": "scrub_step_start",
                "step_id": step.get("id"),
                "category": category,
                "dry_run": step_dry_run,
            },
        )

        try:
            if category == "context":
                human_logger.info("Context: %s", metadata)
            elif category == "msi-uninstall":
                product = metadata.get("product", {})
                product_code = product.get("product_code") or metadata.get("product_code")
                if not product_code:
                    human_logger.warning("Skipping MSI uninstall step without product code: %s", step)
                else:
                    msi_uninstall.uninstall_products([str(product_code)], dry_run=step_dry_run)
            elif category == "c2r-uninstall":
                installation = metadata.get("installation") or metadata
                if not installation:
                    human_logger.warning("Skipping C2R uninstall step without installation metadata")
                else:
                    c2r_uninstall.uninstall_products(installation, dry_run=step_dry_run)
            elif category == "licensing-cleanup":
                metadata["dry_run"] = step_dry_run
                licensing.cleanup_licenses(metadata)
            elif category == "filesystem-cleanup":
                paths = metadata.get("paths", [])
                if not paths:
                    human_logger.info("No filesystem paths supplied; skipping step.")
                else:
                    fs_tools.remove_paths(paths, dry_run=step_dry_run)
            elif category == "registry-cleanup":
                human_logger.info("Registry cleanup step planned for keys: %s", metadata.get("keys", []))
            else:
                human_logger.info("Unhandled plan category %s; skipping.", category)
        except Exception as exc:
            machine_logger.error(
                "scrub_step_failure",
                extra={
                    "event": "scrub_step_failure",
                    "step_id": step.get("id"),
                    "category": category,
                    "error": repr(exc),
                },
            )
            raise
        else:
            machine_logger.info(
                "scrub_step_complete",
                extra={
                    "event": "scrub_step_complete",
                    "step_id": step.get("id"),
                    "category": category,
                    "dry_run": step_dry_run,
                },
            )

    machine_logger.info(
        "scrub_plan_complete",
        extra={"event": "scrub_plan_complete", "step_count": len(steps), "dry_run": global_dry_run},
    )
