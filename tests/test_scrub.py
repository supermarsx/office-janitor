"""!
@brief Scrubber orchestration tests.
@details Validates multi-pass behaviour, dry-run safeguards, and command wiring
against the new OffScrub-based uninstall helpers.
"""

from __future__ import annotations

import pathlib
import sys
from typing import List

PROJECT_ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import logging_ext, scrub


def _context(dry_run: bool = False, options: dict | None = None, pass_index: int = 1) -> dict:
    return {
        "id": "context",
        "category": "context",
        "metadata": {
            "options": options or {},
            "dry_run": dry_run,
            "pass_index": pass_index,
        },
    }


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
        lambda products, dry_run=False: events.append(f"msi:{products[0]['product_code']}:{dry_run}"),
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

    def fake_reprobe(options):
        return {"msi": [], "c2r": [], "filesystem": [], "registry": []}

    monkeypatch.setattr(scrub.detect, "reprobe", fake_reprobe)

    def fake_replan(inventory, options, pass_index=1):
        assert pass_index == 2
        return [
            _context(options.get("dry_run", False), options, pass_index),
            {
                "id": "licensing-2-0",
                "category": "licensing-cleanup",
                "depends_on": ["context"],
                "metadata": {"dry_run": options.get("dry_run", False)},
            },
            {
                "id": "filesystem-2-0",
                "category": "filesystem-cleanup",
                "depends_on": ["licensing-2-0"],
                "metadata": {"paths": [str(tmp_path / "stale")]},
            },
        ]

    monkeypatch.setattr(scrub.plan_module, "build_plan", fake_replan)

    plan = [
        _context(False, {"create_restore_point": True}, 1),
        {
            "id": "msi-1-0",
            "category": "msi-uninstall",
            "metadata": {"product": {"product_code": "{CODE}", "version": "2016"}},
        },
        {
            "id": "c2r-1-0",
            "category": "c2r-uninstall",
            "metadata": {"installation": {"release_ids": ["Test"]}},
        },
        {
            "id": "licensing-1-0",
            "category": "licensing-cleanup",
            "metadata": {},
        },
        {
            "id": "filesystem-1-0",
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
        "msi:{CODE}:False",
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
        lambda products, dry_run=False: recorded.append(f"msi:{dry_run}"),
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

    monkeypatch.setattr(scrub.detect, "reprobe", lambda options: (_ for _ in ()).throw(AssertionError("reprobe should not run")))
    monkeypatch.setattr(scrub.plan_module, "build_plan", lambda inventory, options, pass_index=1: (_ for _ in ()).throw(AssertionError("replan should not run")))

    plan = [
        _context(True, {}, 1),
        {
            "id": "msi-1-0",
            "category": "msi-uninstall",
            "metadata": {"product": {"product_code": "{CODE}", "version": "2016"}},
        },
        {
            "id": "c2r-1-0",
            "category": "c2r-uninstall",
            "metadata": {"installation": {"release_ids": ["Test"]}},
        },
        {
            "id": "licensing-1-0",
            "category": "licensing-cleanup",
            "metadata": {},
        },
        {
            "id": "filesystem-1-0",
            "category": "filesystem-cleanup",
            "metadata": {"paths": ["X"]},
        },
    ]

    scrub.execute_plan(plan, dry_run=True)

    assert recorded == [
        "msi:True",
        "c2r:True",
        "licensing:True",
        "filesystem:True",
    ]


def test_execute_plan_repeats_until_clean(monkeypatch, tmp_path) -> None:
    """!
    @brief Validate multi-pass behaviour with leftover MSI inventory.
    """

    logging_ext.setup_logging(tmp_path)
    events: List[str] = []

    monkeypatch.setattr(scrub.restore_point, "create_restore_point", lambda description: None)
    monkeypatch.setattr(scrub.processes, "terminate_office_processes", lambda names: None)
    monkeypatch.setattr(scrub.tasks_services, "stop_services", lambda services, timeout=30: None)
    monkeypatch.setattr(scrub.tasks_services, "disable_tasks", lambda tasks, dry_run=False: None)

    monkeypatch.setattr(
        scrub.msi_uninstall,
        "uninstall_products",
        lambda products, dry_run=False: events.append(f"msi:{products[0]['product_code']}:{dry_run}"),
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

    inventories = [
        {"msi": [{"product_code": "{LEFTOVER}", "version": "2016"}], "c2r": [], "filesystem": [], "registry": []},
        {"msi": [], "c2r": [], "filesystem": [], "registry": []},
    ]

    def fake_reprobe(options):
        return inventories.pop(0)

    monkeypatch.setattr(scrub.detect, "reprobe", fake_reprobe)

    def fake_replan(inventory, options, pass_index=1):
        if pass_index == 2:
            return [
                _context(options.get("dry_run", False), options, pass_index),
                {
                    "id": "msi-2-0",
                    "category": "msi-uninstall",
                    "metadata": {"product": {"product_code": "{LEFTOVER}", "version": "2016"}},
                },
                {
                    "id": "licensing-2-0",
                    "category": "licensing-cleanup",
                    "depends_on": ["msi-2-0"],
                    "metadata": {},
                },
            ]
        assert pass_index == 3
        return [
            _context(options.get("dry_run", False), options, pass_index),
            {
                "id": "licensing-3-0",
                "category": "licensing-cleanup",
                "depends_on": ["context"],
                "metadata": {},
            },
        ]

    monkeypatch.setattr(scrub.plan_module, "build_plan", fake_replan)

    plan = [
        _context(False, {}, 1),
        {
            "id": "msi-1-0",
            "category": "msi-uninstall",
            "metadata": {"product": {"product_code": "{CODE}", "version": "2016"}},
        },
        {
            "id": "c2r-1-0",
            "category": "c2r-uninstall",
            "metadata": {"installation": {"release_ids": ["Test"]}},
        },
        {
            "id": "licensing-1-0",
            "category": "licensing-cleanup",
            "metadata": {},
        },
    ]

    scrub.execute_plan(plan)

    assert events == [
        "msi:{CODE}:False",
        "c2r:False",
        "msi:{LEFTOVER}:False",
        "licensing:False",
    ]
