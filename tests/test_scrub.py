from __future__ import annotations

import pathlib
import sys
from typing import List

PROJECT_ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import logging_ext, scrub


def test_execute_plan_runs_steps_in_order(monkeypatch, tmp_path) -> None:
    """!
    @brief Ensure scrubber orchestrates restore point, process, and cleanup steps.
    """

    logging_ext.setup_logging(tmp_path)
    events: List[str] = []

    monkeypatch.setattr(
        scrub.restore_point,
        "create_restore_point",
        lambda description: events.append("restore_point"),
    )

    monkeypatch.setattr(
        scrub.processes,
        "terminate_office_processes",
        lambda names: events.append("terminate_processes"),
    )

    monkeypatch.setattr(
        scrub.tasks_services,
        "stop_services",
        lambda services, timeout=30: events.append("stop_services"),
    )

    monkeypatch.setattr(
        scrub.tasks_services,
        "disable_tasks",
        lambda tasks, dry_run=False: events.append("disable_tasks"),
    )

    monkeypatch.setattr(
        scrub.msi_uninstall,
        "uninstall_products",
        lambda codes, dry_run=False: events.append(f"msi:{dry_run}"),
    )

    monkeypatch.setattr(
        scrub.c2r_uninstall,
        "uninstall_products",
        lambda config, dry_run=False: events.append(f"c2r:{dry_run}"),
    )

    monkeypatch.setattr(
        scrub.licensing,
        "cleanup_licenses",
        lambda metadata: events.append(f"licensing:{metadata.get('dry_run')}"),
    )

    monkeypatch.setattr(
        scrub.fs_tools,
        "remove_paths",
        lambda paths, dry_run=False: events.append(f"filesystem:{dry_run}"),
    )

    plan = [
        {
            "id": "context",
            "category": "context",
            "metadata": {"options": {"create_restore_point": True}, "dry_run": False},
        },
        {
            "id": "msi-0",
            "category": "msi-uninstall",
            "metadata": {"product": {"product_code": "{CODE}"}},
        },
        {
            "id": "c2r-0",
            "category": "c2r-uninstall",
            "metadata": {"installation": {"release_ids": ["Test"]}},
        },
        {
            "id": "licensing-0",
            "category": "licensing-cleanup",
            "metadata": {},
        },
        {
            "id": "filesystem-0",
            "category": "filesystem-cleanup",
            "metadata": {"paths": [str(tmp_path / "stale")]},
        },
    ]

    scrub.execute_plan(plan)

    assert events == [
        "restore_point",
        "terminate_processes",
        "stop_services",
        "disable_tasks",
        "msi:False",
        "c2r:False",
        "licensing:False",
        "filesystem:False",
    ]


def test_execute_plan_dry_run_skips_mutations(monkeypatch, tmp_path) -> None:
    """!
    @brief Dry-run should avoid restore/process work but still call uninstall stubs.
    """

    logging_ext.setup_logging(tmp_path)

    monkeypatch.setattr(
        scrub.restore_point,
        "create_restore_point",
        lambda description: (_ for _ in ()).throw(AssertionError("restore point should not run")),
    )
    monkeypatch.setattr(
        scrub.processes,
        "terminate_office_processes",
        lambda names: (_ for _ in ()).throw(AssertionError("terminate should not run")),
    )
    monkeypatch.setattr(
        scrub.tasks_services,
        "stop_services",
        lambda services, timeout=30: (_ for _ in ()).throw(AssertionError("stop services should not run")),
    )
    monkeypatch.setattr(
        scrub.tasks_services,
        "disable_tasks",
        lambda tasks, dry_run=False: (_ for _ in ()).throw(AssertionError("disable tasks should not run")),
    )

    recorded: List[str] = []

    monkeypatch.setattr(
        scrub.msi_uninstall,
        "uninstall_products",
        lambda codes, dry_run=False: recorded.append(f"msi:{dry_run}"),
    )
    monkeypatch.setattr(
        scrub.c2r_uninstall,
        "uninstall_products",
        lambda config, dry_run=False: recorded.append(f"c2r:{dry_run}"),
    )
    monkeypatch.setattr(
        scrub.licensing,
        "cleanup_licenses",
        lambda metadata: recorded.append(f"licensing:{metadata.get('dry_run')}"),
    )
    monkeypatch.setattr(
        scrub.fs_tools,
        "remove_paths",
        lambda paths, dry_run=False: recorded.append(f"filesystem:{dry_run}"),
    )

    plan = [
        {"id": "context", "category": "context", "metadata": {"dry_run": True}},
        {"id": "msi-0", "category": "msi-uninstall", "metadata": {"product": {"product_code": "{CODE}"}}},
        {"id": "c2r-0", "category": "c2r-uninstall", "metadata": {"installation": {"release_ids": ["Test"]}}},
        {"id": "licensing-0", "category": "licensing-cleanup", "metadata": {}},
        {"id": "filesystem-0", "category": "filesystem-cleanup", "metadata": {"paths": ["X"]}},
    ]

    scrub.execute_plan(plan, dry_run=True)

    assert recorded == [
        "msi:True",
        "c2r:True",
        "licensing:True",
        "filesystem:True",
    ]
