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
            "filesystem-cleanup",
            "registry-cleanup",
        ]

        context = plan_steps[0]
        assert context["metadata"]["mode"] == "auto-all"

        licensing = next(step for step in plan_steps if step["category"] == "licensing-cleanup")
        assert set(licensing["depends_on"]) == {"msi-0", "c2r-0"}
        assert licensing["metadata"]["dry_run"] is False

        filesystem = next(step for step in plan_steps if step["category"] == "filesystem-cleanup")
        assert set(filesystem["depends_on"]) == {"msi-0", "c2r-0"}

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

        msi_steps = [step for step in plan_steps if step["category"] == "msi-uninstall"]
        assert len(msi_steps) == 1
        assert msi_steps[0]["metadata"]["version"] == "2016"

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
            "filesystem-cleanup",
            "registry-cleanup",
        }

        context = plan_steps[0]
        assert context["metadata"]["mode"] == "cleanup-only"

        licensing = next(step for step in plan_steps if step["category"] == "licensing-cleanup")
        assert licensing["depends_on"] == ["context"]
        assert licensing["metadata"]["dry_run"] is True

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

