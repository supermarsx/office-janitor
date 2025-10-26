"""!
@brief Planning rule enforcement tests.
@details Validates ordering, filtering, and dependency metadata produced by the
planner when combining inventory signals with CLI options.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Dict, List

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import plan


class TestPlanBuilder:
    """!
    @brief Scenario-driven planner validation.
    @details Each test feeds curated inventory snapshots and CLI selections to
    the planner to ensure resulting plans respect guardrails documented in the
    specification.
    """

    def test_auto_all_generates_ordered_steps(self) -> None:
        """!
        @brief Validate ordering and dependencies for auto-all mode.
        @details Ensures the planner emits MSI then C2R uninstall steps, followed
        by licensing and cleanup actions, all linked with dependency metadata.
        """

        inventory: Dict[str, List[dict]] = {
            "msi": [
                {
                    "product_code": "{91190000-0011-0000-0000-0000000FF1CE}",
                    "display_name": "Microsoft Office Professional Plus 2019",
                    "version": "2019",
                }
            ],
            "c2r": [
                {
                    "release_ids": ["O365ProPlusRetail"],
                    "channel": "Monthly Enterprise Channel",
                    "version": "16.0.17029.20108",
                    "tags": ["365"],
                }
            ],
            "tasks": [
                {"task": r"\\Microsoft\\Office\\TelemetryTask"},
            ],
            "services": [
                {"name": "ClickToRunSvc"},
            ],
            "filesystem": [
                {"path": r"C:\\Program Files\\Microsoft Office"},
            ],
            "registry": [
                {"path": r"HKLM\\SOFTWARE\\Microsoft\\Office\\16.0"},
            ],
        }
        options = {"auto_all": True, "dry_run": False}

        plan_steps = plan.build_plan(inventory, options)

        categories = [step["category"] for step in plan_steps]
        assert categories == [
            "context",
            "msi-uninstall",
            "c2r-uninstall",
            "licensing-cleanup",
            "task-cleanup",
            "service-cleanup",
            "filesystem-cleanup",
            "registry-cleanup",
        ]

        context = plan_steps[0]
        assert context["metadata"]["mode"] == "auto-all"
        assert context["metadata"]["discovered_versions"] == ["2019", "365"]
        assert context["metadata"]["pass_index"] == 1

        licensing = next(step for step in plan_steps if step["category"] == "licensing-cleanup")
        assert set(licensing["depends_on"]) == {"msi-1-0", "c2r-1-0"}
        assert licensing["metadata"]["dry_run"] is False

        task_step = next(step for step in plan_steps if step["category"] == "task-cleanup")
        assert task_step["depends_on"] == ["licensing-1-0"]
        assert task_step["metadata"]["tasks"] == [r"\\Microsoft\\Office\\TelemetryTask"]

        service_step = next(step for step in plan_steps if step["category"] == "service-cleanup")
        assert service_step["depends_on"] == [task_step["id"]]
        assert service_step["metadata"]["services"] == ["ClickToRunSvc"]

        filesystem = next(step for step in plan_steps if step["category"] == "filesystem-cleanup")
        assert filesystem["depends_on"] == [service_step["id"]]

        msi_step = next(step for step in plan_steps if step["category"] == "msi-uninstall")
        c2r_step = next(step for step in plan_steps if step["category"] == "c2r-uninstall")
        assert msi_step["metadata"]["version"] == "2019"
        assert c2r_step["metadata"]["version"] == "365"

    def test_target_mode_filters_inventory(self) -> None:
        """!
        @brief Ensure targeted mode restricts uninstall scope.
        @details Only MSI installations matching the requested version should be
        scheduled when `--target` is supplied.
        """

        inventory: Dict[str, List[dict]] = {
            "msi": [
                {
                    "product_code": "{91160000-0011-0000-0000-0000000FF1CE}",
                    "display_name": "Microsoft Office Professional Plus 2016",
                    "version": "2016",
                },
                {
                    "product_code": "{91190000-0011-0000-0000-0000000FF1CE}",
                    "display_name": "Microsoft Office Professional Plus 2019",
                    "version": "2019",
                },
            ],
        }
        options = {"target": "2016"}

        plan_steps = plan.build_plan(inventory, options)

        msi_steps = [step for step in plan_steps if step["category"] == "msi-uninstall"]
        assert len(msi_steps) == 1
        assert msi_steps[0]["metadata"]["version"] == "2016"

        context = plan_steps[0]
        assert context["metadata"]["target_versions"] == ["2016"]
        assert context["metadata"]["unsupported_targets"] == []
        assert context["metadata"]["discovered_versions"] == ["2016", "2019"]

    def test_target_mode_skips_unknown_versions(self) -> None:
        """!
        @brief Unknown entries are skipped during targeted uninstalls.
        @details When an installation lacks version metadata it must not be
        scheduled for removal during targeted runs.
        """

        inventory: Dict[str, List[dict]] = {
            "msi": [
                {
                    "product_code": "{91160000-0011-0000-0000-0000000FF1CE}",
                    "display_name": "Microsoft Office Professional Plus 2016",
                    "version": "2016",
                },
                {
                    "product_code": "{91190000-0011-0000-0000-0000000FF1CE}",
                    "display_name": "Microsoft Office Professional Plus 2019",
                    "version": "",
                },
            ],
        }
        options = {"target": "2016"}

        plan_steps = plan.build_plan(inventory, options)

        uninstall_versions = [
            step["metadata"].get("version")
            for step in plan_steps
            if step["category"] == "msi-uninstall"
        ]
        assert uninstall_versions == ["2016"]

    def test_explicit_mode_respects_safety_overrides(self) -> None:
        """!
        @brief Safety flags override conflicting explicit modes.
        @details Even if callers pass `mode="auto-all"`, diagnostics or
        cleanup-only selections must take precedence to prevent uninstalls.
        """

        inventory: Dict[str, List[dict]] = {
            "msi": [
                {
                    "product_code": "{91160000-0011-0000-0000-0000000FF1CE}",
                    "display_name": "Microsoft Office Professional Plus 2016",
                    "version": "2016",
                }
            ]
        }

        diagnose_plan = plan.build_plan(inventory, {"mode": "auto-all", "diagnose": True})
        assert [step["category"] for step in diagnose_plan] == ["context"]
        assert diagnose_plan[0]["metadata"]["mode"] == "diagnose"

        cleanup_plan = plan.build_plan(
            inventory,
            {
                "mode": "target:2016",
                "cleanup_only": True,
                "target": "2016",
            },
        )
        categories = {step["category"] for step in cleanup_plan}
        assert "msi-uninstall" not in categories
        assert cleanup_plan[0]["metadata"]["mode"] == "cleanup-only"

    def test_cleanup_only_skips_uninstall(self) -> None:
        """!
        @brief Confirm cleanup-only omits uninstall actions.
        @details Plan should only include licensing and residue cleanup steps while
        still respecting dry-run metadata.
        """

        inventory: Dict[str, List[dict]] = {
            "msi": [
                {
                    "product_code": "{91160000-0011-0000-0000-0000000FF1CE}",
                    "display_name": "Microsoft Office Professional Plus 2016",
                    "version": "2016",
                }
            ],
            "tasks": [
                {"task": r"\\Microsoft\\Office\\TelemetryTask"},
            ],
            "services": [
                {"name": "ClickToRunSvc"},
            ],
            "filesystem": [
                {"path": r"C:\\Program Files\\Microsoft Office"},
            ],
            "registry": [
                {"path": r"HKLM\\SOFTWARE\\Microsoft\\Office\\16.0"},
            ],
        }
        options = {"cleanup_only": True, "dry_run": True, "auto_all": True}

        plan_steps = plan.build_plan(inventory, options)

        categories = {step["category"] for step in plan_steps}
        assert "msi-uninstall" not in categories
        assert "c2r-uninstall" not in categories
        assert {step["category"] for step in plan_steps} == {
            "context",
            "licensing-cleanup",
            "task-cleanup",
            "service-cleanup",
            "filesystem-cleanup",
            "registry-cleanup",
        }

        context = plan_steps[0]
        assert context["metadata"]["mode"] == "cleanup-only"

        licensing = next(step for step in plan_steps if step["category"] == "licensing-cleanup")
        assert licensing["depends_on"] == ["context"]
        assert licensing["metadata"]["dry_run"] is True

        task_step = next(step for step in plan_steps if step["category"] == "task-cleanup")
        assert task_step["depends_on"] == [licensing["id"]]

        service_step = next(step for step in plan_steps if step["category"] == "service-cleanup")
        assert service_step["depends_on"] == [task_step["id"]]

    def test_plan_includes_task_and_service_cleanup(self) -> None:
        """!
        @brief Planner emits task and service cleanup steps when inventory reports them.
        """

        inventory: Dict[str, List[dict]] = {
            "tasks": [
                {"task": r"\\Microsoft\\Office\\TelemetryTask"},
                {"name": r"\\Microsoft\\Office\\OtherTask"},
            ],
            "services": [
                {"name": "ClickToRunSvc"},
                {"service": "ose"},
            ],
        }
        options = {"cleanup_only": True}

        plan_steps = plan.build_plan(inventory, options)

        task_step = next(step for step in plan_steps if step["category"] == "task-cleanup")
        assert task_step["metadata"]["tasks"] == [
            r"\\Microsoft\\Office\\TelemetryTask",
            r"\\Microsoft\\Office\\OtherTask",
        ]

        service_step = next(step for step in plan_steps if step["category"] == "service-cleanup")
        assert service_step["metadata"]["services"] == ["ClickToRunSvc", "ose"]
        assert service_step["depends_on"] == [task_step["id"]]

    def test_diagnose_mode_is_context_only(self) -> None:
        """!
        @brief Diagnostics mode must not contain action steps.
        @details Planner should emit only the context metadata when operating in
        diagnostics mode per the specification.
        """

        inventory: Dict[str, List[dict]] = {
            "msi": [
                {
                    "product_code": "{91160000-0011-0000-0000-0000000FF1CE}",
                    "display_name": "Microsoft Office Professional Plus 2016",
                    "version": "2016",
                }
            ],
        }
        options = {"diagnose": True, "target": "2016", "auto_all": True}

        plan_steps = plan.build_plan(inventory, options)

        assert len(plan_steps) == 1
        assert plan_steps[0]["category"] == "context"
        assert plan_steps[0]["metadata"]["mode"] == "diagnose"

    def test_second_pass_ids_include_pass_index(self) -> None:
        """!
        @brief Subsequent passes use distinct identifiers for uninstall steps.
        """

        inventory: Dict[str, List[dict]] = {
            "msi": [
                {
                    "product_code": "{91160000-0011-0000-0000-0000000FF1CE}",
                    "display_name": "Microsoft Office Professional Plus 2016",
                    "version": "2016",
                }
            ],
            "c2r": [
                {
                    "release_ids": ["O365ProPlusRetail"],
                    "channel": "Monthly Enterprise Channel",
                    "version": "16.0.17029.20108",
                    "tags": ["365"],
                }
            ],
        }
        options = {"auto_all": True}

        plan_steps = plan.build_plan(inventory, options, pass_index=2)

        msi_ids = [step["id"] for step in plan_steps if step["category"] == "msi-uninstall"]
        c2r_ids = [step["id"] for step in plan_steps if step["category"] == "c2r-uninstall"]
        assert msi_ids == ["msi-2-0"]
        assert c2r_ids == ["c2r-2-0"]
        context = plan_steps[0]
        assert context["metadata"]["pass_index"] == 2

