"""!
@brief Scrubber orchestration tests.
@details Validates multi-pass behaviour, dry-run safeguards, and command wiring
against the new OffScrub-based uninstall helpers.
"""

from __future__ import annotations

import json
import pathlib
import sys

PROJECT_ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import logging_ext, scrub  # noqa: E402


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


def test_log_summary_mentions_reboot(tmp_path) -> None:
    """!
    @brief Scrub summary should surface reboot recommendations.
    """

    logging_ext.setup_logging(tmp_path)

    result = scrub.StepResult(
        step_id="context",
        category="context",
        status="success",
        attempts=1,
        dry_run=False,
    )
    result.details["reboot_recommended"] = True
    result.details["reboot_services"] = ["ClickToRunSvc"]
    result.started_at = 0.0
    result.completed_at = 1.0

    scrub._log_summary([result], passes=1, dry_run=False)
    human_log = (tmp_path / "human.log").read_text(encoding="utf-8")
    assert "reboot recommended" in human_log.lower()

    machine_log = (tmp_path / "events.jsonl").read_text(encoding="utf-8").splitlines()
    events = [json.loads(line) for line in machine_log if line.strip()]
    summary = next(event for event in events if event.get("event") == "scrub_summary")
    assert summary["reboot_recommended"] is True
    assert summary["reboot_services"] == ["ClickToRunSvc"]


def test_execute_plan_runs_steps_in_order(monkeypatch, tmp_path) -> None:
    """!
    @brief Ensure scrubber orchestrates restore point, process, and cleanup steps.
    """

    logging_ext.setup_logging(tmp_path)
    events: list[str] = []

    monkeypatch.setattr(
        scrub.restore_point,
        "create_restore_point",
        lambda description, **kwargs: events.append("restore_point"),
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
        scrub.tasks_services,
        "remove_tasks",
        lambda tasks, dry_run=False: events.append(f"remove_tasks:{dry_run}"),
    )
    monkeypatch.setattr(
        scrub.tasks_services,
        "delete_services",
        lambda services, dry_run=False: events.append(f"delete_services:{dry_run}"),
    )

    monkeypatch.setattr(
        scrub.msi_uninstall,
        "uninstall_products",
        lambda products, dry_run=False: events.append(
            f"msi:{products[0]['product_code']}:{dry_run}"
        ),
    )

    monkeypatch.setattr(
        scrub.c2r_uninstall,
        "uninstall_products",
        lambda config, dry_run=False, force=False: events.append(f"c2r:{dry_run}"),
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
        return {
            "msi": [],
            "c2r": [],
            "tasks": [],
            "services": [],
            "filesystem": [],
            "registry": [],
        }

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
                "id": "tasks-2-0",
                "category": "task-cleanup",
                "depends_on": ["licensing-2-0"],
                "metadata": {"tasks": [r"\\Microsoft\\Office\\TelemetryTask"]},
            },
            {
                "id": "services-2-0",
                "category": "service-cleanup",
                "depends_on": ["tasks-2-0"],
                "metadata": {"services": ["ClickToRunSvc"]},
            },
            {
                "id": "filesystem-2-0",
                "category": "filesystem-cleanup",
                "depends_on": ["services-2-0"],
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
            "id": "tasks-1-0",
            "category": "task-cleanup",
            "metadata": {"tasks": [r"\\Microsoft\\Office\\TelemetryTask"]},
        },
        {
            "id": "services-1-0",
            "category": "service-cleanup",
            "metadata": {"services": ["ClickToRunSvc"]},
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
        "remove_tasks:False",
        "delete_services:False",
        "filesystem:False",
    ]


def test_execute_plan_dry_run_skips_mutations(monkeypatch, tmp_path) -> None:
    """!
    @brief Dry-run should avoid restore/process work but still call uninstall stubs.
    """

    logging_ext.setup_logging(tmp_path)

    restore_calls: list[tuple[str, bool]] = []

    def fake_restore(description: str, *, dry_run: bool = False) -> bool:
        restore_calls.append((description, dry_run))
        return dry_run

    monkeypatch.setattr(scrub.restore_point, "create_restore_point", fake_restore)
    monkeypatch.setattr(
        scrub.processes,
        "terminate_office_processes",
        lambda names: (_ for _ in ()).throw(AssertionError("terminate should not run")),
    )
    monkeypatch.setattr(
        scrub.tasks_services,
        "stop_services",
        lambda services, timeout=30: (_ for _ in ()).throw(
            AssertionError("stop services should not run")
        ),
    )
    monkeypatch.setattr(
        scrub.tasks_services,
        "disable_tasks",
        lambda tasks, dry_run=False: (_ for _ in ()).throw(
            AssertionError("disable tasks should not run")
        ),
    )
    monkeypatch.setattr(
        scrub.tasks_services,
        "remove_tasks",
        lambda tasks, dry_run=False: recorded.append(f"remove_tasks:{dry_run}"),
    )
    monkeypatch.setattr(
        scrub.tasks_services,
        "delete_services",
        lambda services, dry_run=False: recorded.append(f"delete_services:{dry_run}"),
    )

    recorded: list[str] = []

    monkeypatch.setattr(
        scrub.msi_uninstall,
        "uninstall_products",
        lambda products, dry_run=False: recorded.append(f"msi:{dry_run}"),
    )
    monkeypatch.setattr(
        scrub.c2r_uninstall,
        "uninstall_products",
        lambda config, dry_run=False, force=False: recorded.append(f"c2r:{dry_run}"),
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

    monkeypatch.setattr(
        scrub.detect,
        "reprobe",
        lambda options: (_ for _ in ()).throw(AssertionError("reprobe should not run")),
    )
    monkeypatch.setattr(
        scrub.plan_module,
        "build_plan",
        lambda inventory, options, pass_index=1: (_ for _ in ()).throw(
            AssertionError("replan should not run")
        ),
    )

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
            "id": "tasks-1-0",
            "category": "task-cleanup",
            "metadata": {"tasks": [r"\\Microsoft\\Office\\TelemetryTask"]},
        },
        {
            "id": "services-1-0",
            "category": "service-cleanup",
            "metadata": {"services": ["ClickToRunSvc"]},
        },
        {
            "id": "filesystem-1-0",
            "category": "filesystem-cleanup",
            "metadata": {"paths": ["X"]},
        },
    ]

    scrub.execute_plan(plan, dry_run=True)

    assert restore_calls == []
    assert recorded == [
        "msi:True",
        "c2r:True",
        "licensing:True",
        "remove_tasks:True",
        "delete_services:True",
        "filesystem:True",
    ]


def test_execute_plan_repeats_until_clean(monkeypatch, tmp_path) -> None:
    """!
    @brief Validate multi-pass behaviour with leftover MSI inventory.
    """

    logging_ext.setup_logging(tmp_path)
    events: list[str] = []

    monkeypatch.setattr(
        scrub.restore_point,
        "create_restore_point",
        lambda description, **kwargs: None,
    )
    monkeypatch.setattr(scrub.processes, "terminate_office_processes", lambda names: None)
    monkeypatch.setattr(scrub.tasks_services, "stop_services", lambda services, timeout=30: None)
    monkeypatch.setattr(scrub.tasks_services, "disable_tasks", lambda tasks, dry_run=False: None)

    monkeypatch.setattr(
        scrub.msi_uninstall,
        "uninstall_products",
        lambda products, dry_run=False: events.append(
            f"msi:{products[0]['product_code']}:{dry_run}"
        ),
    )
    monkeypatch.setattr(
        scrub.c2r_uninstall,
        "uninstall_products",
        lambda config, dry_run=False, force=False: events.append(f"c2r:{dry_run}"),
    )
    monkeypatch.setattr(
        scrub.licensing,
        "cleanup_licenses",
        lambda metadata: events.append(f"licensing:{metadata.get('dry_run')}"),
    )

    inventories = [
        {
            "msi": [{"product_code": "{LEFTOVER}", "version": "2016"}],
            "c2r": [],
            "filesystem": [],
            "registry": [],
        },
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

    # Note: max_passes must be explicitly set to enable multi-pass behavior (default is 1)
    plan = [
        _context(False, {"max_passes": 3}, 1),
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


def test_registry_cleanup_exports_and_deletes(monkeypatch, tmp_path) -> None:
    """!
    @brief Ensure registry cleanup exports keys before deletion.
    """

    logdir = tmp_path / "logs"
    logging_ext.setup_logging(logdir)

    backup_dir = tmp_path / "backups"
    recorded: dict[str, tuple[list[str], object]] = {}

    def fake_export(keys, destination):
        recorded["export"] = (list(keys), destination)

    def fake_delete(keys, dry_run=False):
        recorded["delete"] = (list(keys), dry_run)

    monkeypatch.setattr(scrub.registry_tools, "export_keys", fake_export)
    monkeypatch.setattr(scrub.registry_tools, "delete_keys", fake_delete)

    plan = [
        {
            "id": "context",
            "category": "context",
            "metadata": {
                "options": {"backup": str(backup_dir), "logdir": str(logdir)},
                "dry_run": False,
                "pass_index": 1,
                "backup_destination": str(backup_dir),
                "log_directory": str(logdir),
            },
        },
        {
            "id": "registry-1-0",
            "category": "registry-cleanup",
            "metadata": {
                "keys": ["HKLM\\Software\\Test"],
                "dry_run": False,
                "backup_destination": str(backup_dir),
                "log_directory": str(logdir),
            },
        },
    ]

    scrub._execute_steps(plan, scrub.CLEANUP_CATEGORIES, False)

    assert recorded["export"] == (["HKLM\\Software\\Test"], str(backup_dir))
    assert recorded["delete"] == (["HKLM\\Software\\Test"], False)


def test_filesystem_cleanup_preserves_templates(monkeypatch, tmp_path) -> None:
    """!
    @brief Ensure filesystem cleanup skips user templates when preservation is requested.
    """

    logging_ext.setup_logging(tmp_path)

    removed: list[list[str]] = []

    monkeypatch.setattr(
        scrub.fs_tools,
        "remove_paths",
        lambda paths, dry_run=False: removed.append(list(paths)),
    )

    context_metadata = {"options": {"keep_templates": True}}
    metadata = {
        "paths": [
            r"C:\\Users\\User\\AppData\\Roaming\\Microsoft\\Templates",
            r"C:\\ProgramData\\Microsoft\\Office",
        ],
        "preserve_templates": True,
    }

    scrub._perform_filesystem_cleanup(metadata, context_metadata, dry_run=False)

    assert removed == [[r"C:\\ProgramData\\Microsoft\\Office"]]


def test_registry_cleanup_generates_backup_when_missing(monkeypatch, tmp_path) -> None:
    """!
    @brief Missing backup destinations should fall back to the log directory.
    """

    logging_ext.setup_logging(tmp_path)

    recorded: dict[str, object] = {}

    def fake_export(keys, destination):
        recorded["export"] = {"keys": list(keys), "destination": destination}

    def fake_delete(keys, dry_run=False):
        recorded["delete"] = {"keys": list(keys), "dry_run": dry_run}

    monkeypatch.setattr(scrub.registry_tools, "export_keys", fake_export)
    monkeypatch.setattr(scrub.registry_tools, "delete_keys", fake_delete)

    metadata = {"keys": ["HKLM\\Software\\Test"], "log_directory": str(tmp_path)}

    scrub._perform_registry_cleanup(
        metadata,
        dry_run=False,
        default_backup=None,
        default_logdir=str(tmp_path),
    )

    assert recorded["export"]["keys"] == ["HKLM\\Software\\Test"]
    assert pathlib.Path(recorded["export"]["destination"]).parent == tmp_path
    assert recorded["delete"] == {"keys": ["HKLM\\Software\\Test"], "dry_run": False}


def test_registry_cleanup_sorts_child_paths_first(monkeypatch, tmp_path) -> None:
    """!
    @brief Registry cleanup deletes nested keys before parents to avoid errors.
    """

    logging_ext.setup_logging(tmp_path)

    recorded: dict[str, list[str]] = {}

    def fake_delete(keys, dry_run=False):
        recorded["keys"] = list(keys)

    monkeypatch.setattr(scrub.registry_tools, "delete_keys", fake_delete)

    metadata = {
        "keys": [
            "HKLM\\SOFTWARE\\Microsoft\\Office",
            "HKLM\\SOFTWARE\\Microsoft\\Office\\16.0",
        ]
    }

    scrub._perform_registry_cleanup(
        metadata,
        dry_run=True,
        default_backup=None,
        default_logdir=None,
    )

    assert recorded["keys"] == [
        "HKLM\\SOFTWARE\\Microsoft\\Office\\16.0",
        "HKLM\\SOFTWARE\\Microsoft\\Office",
    ]
