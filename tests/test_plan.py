"""!
@brief Planning rule enforcement tests.
@details Validates ordering, filtering, and dependency metadata produced by the
planner when combining inventory signals with CLI options.
"""

from __future__ import annotations

import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import constants, plan  # noqa: E402


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

        inventory: dict[str, list[dict]] = {
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
        assert categories[:2] == ["context", "detect"]

        c2r_indices = [
            index for index, category in enumerate(categories) if category == "c2r-uninstall"
        ]
        assert c2r_indices and min(c2r_indices) == 2
        first_msi_index = categories.index("msi-uninstall")
        assert max(c2r_indices) < first_msi_index

        trailing_categories = categories[first_msi_index + 1 :]
        for expected in (
            "licensing-cleanup",
            "task-cleanup",
            "service-cleanup",
            "filesystem-cleanup",
            "registry-cleanup",
        ):
            assert expected in trailing_categories

        context = plan_steps[0]
        assert context["metadata"]["mode"] == "auto-all"
        assert context["metadata"]["discovered_versions"] == ["2019", "365"]
        assert context["metadata"]["pass_index"] == 1
        summary = context["metadata"]["summary"]
        assert summary["total_steps"] == len(plan_steps)
        assert summary["categories"]["detect"] == 1

        detect_step = plan_steps[1]
        assert detect_step["category"] == "detect"
        assert detect_step["depends_on"] == ["context"]

        licensing = next(step for step in plan_steps if step["category"] == "licensing-cleanup")
        c2r_step_ids = [step["id"] for step in plan_steps if step["category"] == "c2r-uninstall"]
        assert set(licensing["depends_on"]) == {"msi-1-0", *c2r_step_ids}
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
        detection_c2r_step = next(
            step
            for step in plan_steps
            if step["category"] == "c2r-uninstall"
            and "O365ProPlusRetail" in step["metadata"]["installation"].get("release_ids", [])
        )
        assert msi_step["metadata"]["version"] == "2019"
        assert detection_c2r_step["metadata"]["version"] == "365"

    def test_auto_all_seeds_default_c2r_inventory(self) -> None:
        """!
        @brief Auto-all mode should seed Click-to-Run uninstall steps by default.
        @details When detection yields no C2R entries the planner should still
        schedule curated release identifiers for modern Office suites so the
        uninstall sequence mirrors OffScrub automation.
        """

        inventory: dict[str, list[dict]] = {}
        options = {"auto_all": True}

        plan_steps = plan.build_plan(inventory, options)

        c2r_steps = [step for step in plan_steps if step["category"] == "c2r-uninstall"]
        assert c2r_steps, "seeded plan should include Click-to-Run uninstall steps"

        optional_families = {"project", "visio", "onenote"}
        expected_release_ids = [
            release_id
            for release_id, metadata in constants.DEFAULT_AUTO_ALL_C2R_RELEASES.items()
            if str(
                metadata.get("family")
                or constants.C2R_PRODUCT_RELEASES.get(release_id, {}).get("family")
                or "office"
            ).lower()
            not in optional_families
        ]
        seeded_release_ids = [
            release_id
            for step in c2r_steps
            for release_id in step["metadata"]["installation"].get("release_ids", [])
        ]
        assert seeded_release_ids == expected_release_ids

        def _expected_version(release_id: str) -> str:
            details = constants.DEFAULT_AUTO_ALL_C2R_RELEASES[release_id]
            candidate = str(details.get("default_version") or "").strip()
            if candidate:
                return candidate
            base = constants.C2R_PRODUCT_RELEASES.get(release_id, {})
            versions = base.get("supported_versions") or ()
            if versions:
                return str(list(versions)[-1])
            return "c2r"

        expected_versions = {_expected_version(release_id) for release_id in expected_release_ids}

        context = plan_steps[0]
        assert context["metadata"]["mode"] == "auto-all"
        assert context["metadata"]["discovered_versions"] == []
        assert context["metadata"]["target_versions"] == sorted(expected_versions)

    def test_auto_all_seeding_skips_detected_release_ids(self) -> None:
        """!
        @brief Detected Click-to-Run suites should not be duplicated when seeded.
        """

        inventory: dict[str, list[dict]] = {
            "c2r": [
                {
                    "product": "Microsoft 365 Apps for enterprise",
                    "version": "365",
                    "release_ids": ["O365ProPlusRetail"],
                    "channel": "Current Channel",
                }
            ]
        }
        options = {"auto_all": True}

        plan_steps = plan.build_plan(inventory, options)

        c2r_release_ids = [
            release_id
            for step in plan_steps
            if step["category"] == "c2r-uninstall"
            for release_id in step["metadata"]["installation"].get("release_ids", [])
        ]
        assert c2r_release_ids.count("O365ProPlusRetail") == 1

    def test_auto_all_include_components_adds_optional_c2r(self) -> None:
        """!
        @brief Optional components should be seeded when explicitly included.
        """

        inventory: dict[str, list[dict]] = {}
        options = {"auto_all": True, "include": "Visio"}

        plan_steps = plan.build_plan(inventory, options)

        c2r_release_ids = [
            release_id
            for step in plan_steps
            if step["category"] == "c2r-uninstall"
            for release_id in step["metadata"]["installation"].get("release_ids", [])
        ]

        assert "VisioProRetail" in c2r_release_ids

    def test_uninstall_order_matches_offscrub_sequence(self) -> None:
        """!
        @brief Ensure uninstall ordering mirrors the reference script.
        @details Click-to-Run removals must precede MSI generations ordered as
        2016+, 2013, 2010, then 2007.
        """

        inventory: dict[str, list[dict]] = {
            "msi": [
                {"product_code": "{A}", "display_name": "Office 2007", "version": "2007"},
                {"product_code": "{B}", "display_name": "Office 2013", "version": "2013"},
                {"product_code": "{C}", "display_name": "Office 2010", "version": "2010"},
                {"product_code": "{D}", "display_name": "Office 2016", "version": "2016"},
            ],
            "c2r": [
                {
                    "release_ids": ["O365ProPlusRetail"],
                    "channel": "Current Channel",
                    "version": "365",
                }
            ],
        }
        plan_steps = plan.build_plan(inventory, {"auto_all": True})

        categories = [step["category"] for step in plan_steps]
        c2r_indices = [
            index for index, category in enumerate(categories) if category == "c2r-uninstall"
        ]
        msi_indices = [
            index for index, category in enumerate(categories) if category == "msi-uninstall"
        ]

        assert c2r_indices and msi_indices
        assert max(c2r_indices) < min(msi_indices)

        msi_versions = [
            step["metadata"].get("version")
            for step in plan_steps
            if step["category"] == "msi-uninstall"
        ]

        assert msi_versions == ["2016", "2013", "2010", "2007"]

    def test_msi_display_versions_use_supported_priority(self) -> None:
        """!
        @brief Ensure MSI sorting uses supported version metadata for display builds.
        @details DisplayVersion values such as ``16.0.10386`` should still map to the
        OffScrub 2016+ stage based on ``supported_versions`` metadata.
        """

        inventory: dict[str, list[dict]] = {
            "msi": [
                {
                    "product_code": "{C}",
                    "display_name": "Office 2010",
                    "version": "14.0.7237.5000",
                    "properties": {"supported_versions": ["2010"]},
                },
                {
                    "product_code": "{A}",
                    "display_name": "Office 2016",
                    "version": "16.0.10386.20017",
                    "properties": {"supported_versions": ["2016", "2019"]},
                },
                {
                    "product_code": "{B}",
                    "display_name": "Office 2013",
                    "version": "15.0.5189.1000",
                    "properties": {"supported_versions": ["2013"]},
                },
            ]
        }

        plan_steps = plan.build_plan(inventory, {"auto_all": True})

        uninstall_versions = [
            step["metadata"].get("version")
            for step in plan_steps
            if step["category"] == "msi-uninstall"
        ]

        assert uninstall_versions == [
            "16.0.10386.20017",
            "15.0.5189.1000",
            "14.0.7237.5000",
        ]

    def test_target_mode_filters_inventory(self) -> None:
        """!
        @brief Ensure targeted mode restricts uninstall scope.
        @details Only MSI installations matching the requested version should be
        scheduled when `--target` is supplied.
        """

        inventory: dict[str, list[dict]] = {
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
        assert msi_steps[0]["depends_on"] == ["detect-1-0"]

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

        inventory: dict[str, list[dict]] = {
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

        inventory: dict[str, list[dict]] = {
            "msi": [
                {
                    "product_code": "{91160000-0011-0000-0000-0000000FF1CE}",
                    "display_name": "Microsoft Office Professional Plus 2016",
                    "version": "2016",
                }
            ]
        }

        diagnose_plan = plan.build_plan(inventory, {"mode": "auto-all", "diagnose": True})
        assert [step["category"] for step in diagnose_plan] == ["context", "detect"]
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

        inventory: dict[str, list[dict]] = {
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
            "detect",
            "licensing-cleanup",
            "task-cleanup",
            "service-cleanup",
            "filesystem-cleanup",
            "registry-cleanup",
        }

        context = plan_steps[0]
        assert context["metadata"]["mode"] == "cleanup-only"

        licensing = next(step for step in plan_steps if step["category"] == "licensing-cleanup")
        assert licensing["depends_on"] == ["detect-1-0"]
        assert licensing["metadata"]["dry_run"] is True

        task_step = next(step for step in plan_steps if step["category"] == "task-cleanup")
        assert task_step["depends_on"] == [licensing["id"]]

        service_step = next(step for step in plan_steps if step["category"] == "service-cleanup")
        assert service_step["depends_on"] == [task_step["id"]]

    def test_plan_includes_task_and_service_cleanup(self) -> None:
        """!
        @brief Planner emits task and service cleanup steps when inventory reports them.
        """

        inventory: dict[str, list[dict]] = {
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
        detect_step = next(step for step in plan_steps if step["category"] == "detect")
        assert detect_step["depends_on"] == ["context"]

    def test_diagnose_mode_is_context_only(self) -> None:
        """!
        @brief Diagnostics mode must not contain action steps.
        @details Planner should emit only the context metadata when operating in
        diagnostics mode per the specification.
        """

        inventory: dict[str, list[dict]] = {
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

        assert [step["category"] for step in plan_steps] == ["context", "detect"]
        assert plan_steps[0]["metadata"]["mode"] == "diagnose"

    def test_second_pass_ids_include_pass_index(self) -> None:
        """!
        @brief Subsequent passes use distinct identifiers for uninstall steps.
        """

        inventory: dict[str, list[dict]] = {
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

        optional_families = {"project", "visio", "onenote"}
        expected_release_ids = [
            release_id
            for release_id, metadata in constants.DEFAULT_AUTO_ALL_C2R_RELEASES.items()
            if str(
                metadata.get("family")
                or constants.C2R_PRODUCT_RELEASES.get(release_id, {}).get("family")
                or "office"
            ).lower()
            not in optional_families
        ]
        assert len(c2r_ids) == len(expected_release_ids)
        assert c2r_ids[0] == "c2r-2-0"
        assert all(identifier.startswith("c2r-2-") for identifier in c2r_ids)
        context = plan_steps[0]
        assert context["metadata"]["pass_index"] == 2
        detect_ids = [step["id"] for step in plan_steps if step["category"] == "detect"]
        assert detect_ids == ["detect-2-0"]

    def test_include_components_recorded_in_context(self) -> None:
        """!
        @brief Include flags should be normalized and stored in metadata.
        """

        inventory: dict[str, list[dict]] = {}
        options = {"auto_all": True, "include": "Visio,Project,unknown"}

        plan_steps = plan.build_plan(inventory, options)

        context = plan_steps[0]
        metadata = context["metadata"]
        assert metadata["requested_components"] == ["visio", "project"]
        assert metadata["unsupported_components"] == ["unknown"]
        summary = metadata["summary"]
        assert summary["requested_components"] == ["visio", "project"]
        assert summary["unsupported_components"] == ["unknown"]

    def test_plan_summary_helper(self) -> None:
        """!
        @brief :func:`plan.summarize_plan` aggregates categories and versions.
        """

        inventory: dict[str, list[dict]] = {
            "msi": [
                {
                    "product_code": "{91160000-0011-0000-0000-0000000FF1CE}",
                    "display_name": "Microsoft Office Professional Plus 2016",
                    "version": "2016",
                }
            ],
            "filesystem": [{"path": r"C:\\Office"}],
        }
        options = {"auto_all": True}

        plan_steps = plan.build_plan(inventory, options)
        summary = plan.summarize_plan(plan_steps)

        assert summary["total_steps"] == len(plan_steps)
        assert summary["categories"]["detect"] == 1
        assert summary["uninstall_versions"] == ["2016", "2019", "2021", "2024", "365"]
        assert "filesystem-cleanup" in summary["cleanup_categories"]
        assert summary["actionable_steps"] == len(plan_steps) - 2  # minus context + detect

    def test_dry_run_metadata_propagates_to_steps(self) -> None:
        """!
        @brief Dry-run flag should reach every actionable step emitted by the planner.
        """

        inventory: dict[str, list[dict]] = {
            "msi": [
                {
                    "product_code": "{91160000-0011-0000-0000-0000000FF1CE}",
                    "display_name": "Microsoft Office Professional Plus 2016",
                    "version": "2016",
                }
            ],
            "services": [{"name": "ClickToRunSvc"}],
            "filesystem": [{"path": r"C:\\Program Files\\Microsoft Office"}],
        }
        options = {"auto_all": True, "dry_run": True}

        plan_steps = plan.build_plan(inventory, options)

        assert plan_steps[0]["metadata"]["dry_run"] is True
        for step in plan_steps[1:]:
            metadata = step.get("metadata", {})
            assert metadata.get("dry_run") is True
            if step["category"] in {"msi-uninstall", "service-cleanup"}:
                assert metadata["dry_run"] is True

    def test_keep_license_option_skips_licensing_cleanup(self) -> None:
        """!
        @brief Licensing cleanup is omitted when keep_license/no_license is set.
        """

        inventory: dict[str, list[dict]] = {"msi": [], "c2r": [], "filesystem": [], "registry": []}
        options = {"auto_all": True, "keep_license": True}

        plan_steps = plan.build_plan(inventory, options)
        categories = {step["category"] for step in plan_steps}
        assert "licensing-cleanup" not in categories


class TestSkipFlagsIntegration:
    """!
    @brief Validate skip flags properly exclude cleanup steps from the plan.
    """

    def test_skip_tasks_excludes_task_cleanup_step(self) -> None:
        """!
        @brief skip_tasks=True should omit task-cleanup from plan even with tasks in inventory.
        """
        inventory: dict[str, list[dict]] = {
            "tasks": [{"task": r"\\Microsoft\\Office\\TelemetryTask"}],
            "filesystem": [],
        }
        options = {"auto_all": True, "skip_tasks": True}

        plan_steps = plan.build_plan(inventory, options)
        categories = {step["category"] for step in plan_steps}
        assert "task-cleanup" not in categories

    def test_skip_services_excludes_service_cleanup_step(self) -> None:
        """!
        @brief skip_services=True should omit service-cleanup from plan.
        """
        inventory: dict[str, list[dict]] = {
            "services": [{"name": "ClickToRunSvc"}],
            "filesystem": [],
        }
        options = {"auto_all": True, "skip_services": True}

        plan_steps = plan.build_plan(inventory, options)
        categories = {step["category"] for step in plan_steps}
        assert "service-cleanup" not in categories

    def test_skip_filesystem_excludes_filesystem_cleanup_step(self) -> None:
        """!
        @brief skip_filesystem=True should omit filesystem-cleanup from plan.
        """
        inventory: dict[str, list[dict]] = {
            "filesystem": [{"path": r"C:\\Program Files\\Microsoft Office"}],
        }
        options = {"auto_all": True, "skip_filesystem": True}

        plan_steps = plan.build_plan(inventory, options)
        categories = {step["category"] for step in plan_steps}
        assert "filesystem-cleanup" not in categories

    def test_skip_registry_excludes_registry_cleanup_step(self) -> None:
        """!
        @brief skip_registry=True should omit registry-cleanup from plan.
        """
        inventory: dict[str, list[dict]] = {
            "registry": [{"path": r"HKLM\\SOFTWARE\\Microsoft\\Office"}],
        }
        options = {"auto_all": True, "skip_registry": True}

        plan_steps = plan.build_plan(inventory, options)
        categories = {step["category"] for step in plan_steps}
        assert "registry-cleanup" not in categories

    def test_multiple_skip_flags_combine(self) -> None:
        """!
        @brief Multiple skip flags should all be honored simultaneously.
        """
        inventory: dict[str, list[dict]] = {
            "tasks": [{"task": r"\\Office\\Task"}],
            "services": [{"name": "OfficeSvc"}],
            "filesystem": [{"path": r"C:\\Office"}],
            "registry": [{"path": r"HKLM\\Office"}],
        }
        options = {
            "auto_all": True,
            "skip_tasks": True,
            "skip_services": True,
            "skip_filesystem": True,
            "skip_registry": True,
        }

        plan_steps = plan.build_plan(inventory, options)
        categories = {step["category"] for step in plan_steps}
        assert "task-cleanup" not in categories
        assert "service-cleanup" not in categories
        assert "filesystem-cleanup" not in categories
        assert "registry-cleanup" not in categories


class TestUninstallMethodFiltering:
    """!
    @brief Validate uninstall_method option filters MSI vs C2R steps.
    """

    def test_uninstall_method_msi_only(self) -> None:
        """!
        @brief uninstall_method='msi' should exclude C2R steps.
        """
        inventory: dict[str, list[dict]] = {
            "msi": [{"product_code": "{12345}", "display_name": "Office 2019", "version": "2019"}],
            "c2r": [{"release_ids": ["O365ProPlusRetail"], "version": "365"}],
        }
        options = {"auto_all": True, "uninstall_method": "msi"}

        plan_steps = plan.build_plan(inventory, options)
        categories = [step["category"] for step in plan_steps]
        assert "msi-uninstall" in categories
        assert "c2r-uninstall" not in categories

    def test_uninstall_method_c2r_only(self) -> None:
        """!
        @brief uninstall_method='c2r' should exclude MSI steps.
        """
        inventory: dict[str, list[dict]] = {
            "msi": [{"product_code": "{12345}", "display_name": "Office 2019", "version": "2019"}],
            "c2r": [{"release_ids": ["O365ProPlusRetail"], "version": "365"}],
        }
        options = {"auto_all": True, "uninstall_method": "c2r"}

        plan_steps = plan.build_plan(inventory, options)
        categories = [step["category"] for step in plan_steps]
        assert "msi-uninstall" not in categories
        assert "c2r-uninstall" in categories

    def test_uninstall_method_auto_includes_both(self) -> None:
        """!
        @brief uninstall_method='auto' (default) should include both MSI and C2R.
        """
        inventory: dict[str, list[dict]] = {
            "msi": [{"product_code": "{12345}", "display_name": "Office 2019", "version": "2019"}],
            "c2r": [{"release_ids": ["O365ProPlusRetail"], "version": "365"}],
        }
        options = {"auto_all": True, "uninstall_method": "auto"}

        plan_steps = plan.build_plan(inventory, options)
        categories = [step["category"] for step in plan_steps]
        assert "msi-uninstall" in categories
        assert "c2r-uninstall" in categories


class TestExtendedCleanupOptions:
    """!
    @brief Validate extended cleanup flags are passed through to step metadata.
    """

    def test_clean_msocache_in_filesystem_metadata(self) -> None:
        """!
        @brief clean_msocache flag should be set in filesystem-cleanup metadata.
        """
        inventory: dict[str, list[dict]] = {}
        options = {"auto_all": True, "clean_msocache": True}

        plan_steps = plan.build_plan(inventory, options)
        fs_step = next((s for s in plan_steps if s["category"] == "filesystem-cleanup"), None)
        assert fs_step is not None
        assert fs_step["metadata"]["clean_msocache"] is True

    def test_clean_appx_in_filesystem_metadata(self) -> None:
        """!
        @brief clean_appx flag should be set in filesystem-cleanup metadata.
        """
        inventory: dict[str, list[dict]] = {}
        options = {"auto_all": True, "clean_appx": True}

        plan_steps = plan.build_plan(inventory, options)
        fs_step = next((s for s in plan_steps if s["category"] == "filesystem-cleanup"), None)
        assert fs_step is not None
        assert fs_step["metadata"]["clean_appx"] is True

    def test_license_granularity_flags_in_metadata(self) -> None:
        """!
        @brief License cleanup flags should appear in licensing-cleanup metadata.
        """
        inventory: dict[str, list[dict]] = {}
        options = {
            "auto_all": True,
            "clean_spp": True,
            "clean_ospp": True,
            "clean_vnext": True,
        }

        plan_steps = plan.build_plan(inventory, options)
        lic_step = next((s for s in plan_steps if s["category"] == "licensing-cleanup"), None)
        assert lic_step is not None
        assert lic_step["metadata"]["clean_spp"] is True
        assert lic_step["metadata"]["clean_ospp"] is True
        assert lic_step["metadata"]["clean_vnext"] is True

    def test_registry_cleanup_flags_in_metadata(self) -> None:
        """!
        @brief Extended registry cleanup flags should appear in registry-cleanup metadata.
        """
        inventory: dict[str, list[dict]] = {}
        options = {
            "auto_all": True,
            "clean_wi_metadata": True,
            "remove_vba": True,
            "clean_com_registry": True,
        }

        plan_steps = plan.build_plan(inventory, options)
        reg_step = next((s for s in plan_steps if s["category"] == "registry-cleanup"), None)
        assert reg_step is not None
        assert reg_step["metadata"]["clean_wi_metadata"] is True
        assert reg_step["metadata"]["remove_vba"] is True
        assert reg_step["metadata"]["clean_com_registry"] is True

    def test_uninstall_entries_included_in_registry_cleanup(self) -> None:
        """!
        @brief Detected uninstall entries should be included in registry cleanup keys for nuclear mode.
        @details The VBS scrubber explicitly removes Control Panel uninstall entries
        like ``HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{guid}``.
        Our planner should include these in the registry-cleanup step when using nuclear scrub level.
        """
        inventory: dict[str, list[dict]] = {
            "uninstall_entries": [
                {
                    "display_name": "Microsoft Office Professional Plus 2016",
                    "registry_handle": (
                        r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
                        r"\{90160000-0011-0000-0000-0000000FF1CE}"
                    ),
                },
                {
                    "display_name": "Microsoft Office 365 ProPlus",
                    "registry_handle": (
                        r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
                        r"\O365ProPlusRetail - en-us"
                    ),
                },
            ],
            "registry": [
                {"path": r"HKLM\SOFTWARE\Microsoft\Office\16.0"},
            ],
        }
        options = {"auto_all": True, "scrub_level": "nuclear"}

        plan_steps = plan.build_plan(inventory, options)
        reg_step = next((s for s in plan_steps if s["category"] == "registry-cleanup"), None)
        assert reg_step is not None
        keys = reg_step["metadata"]["keys"]
        # Should include both detected registry residue and uninstall entries
        assert r"HKLM\SOFTWARE\Microsoft\Office\16.0" in keys
        assert (
            r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{90160000-0011-0000-0000-0000000FF1CE}"
            in keys
        )
        assert (
            r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365ProPlusRetail - en-us"
            in keys
        )

    def test_retry_options_in_uninstall_metadata(self) -> None:
        """!
        @brief Retry options should be propagated to uninstall step metadata.
        """
        inventory: dict[str, list[dict]] = {
            "msi": [{"product_code": "{12345}", "display_name": "Office 2019", "version": "2019"}],
        }
        options = {"auto_all": True, "retries": 5, "retry_delay": 10}

        plan_steps = plan.build_plan(inventory, options)
        msi_step = next((s for s in plan_steps if s["category"] == "msi-uninstall"), None)
        assert msi_step is not None
        assert msi_step["metadata"]["retries"] == 5
        assert msi_step["metadata"]["retry_delay"] == 10


class TestScrubLevelBehavior:
    """!
    @brief Validate scrub_level controls cleanup intensity.
    """

    def test_scrub_level_minimal_skips_all_cleanup(self) -> None:
        """!
        @brief scrub_level='minimal' should skip all cleanup steps.
        """
        inventory: dict[str, list[dict]] = {
            "tasks": [{"task": r"\\Office\\Task"}],
            "services": [{"name": "OfficeSvc"}],
            "filesystem": [{"path": r"C:\\Office"}],
            "registry": [{"path": r"HKLM\\Office"}],
        }
        options = {"auto_all": True, "scrub_level": "minimal"}

        plan_steps = plan.build_plan(inventory, options)
        categories = {step["category"] for step in plan_steps}
        # Minimal should skip all cleanup
        assert "task-cleanup" not in categories
        assert "service-cleanup" not in categories
        assert "filesystem-cleanup" not in categories
        assert "registry-cleanup" not in categories

    def test_scrub_level_standard_includes_detected_residue(self) -> None:
        """!
        @brief scrub_level='standard' (default) includes detected cleanup items.
        """
        inventory: dict[str, list[dict]] = {
            "tasks": [{"task": r"\\Office\\Task"}],
            "filesystem": [{"path": r"C:\\Office"}],
        }
        options = {"auto_all": True, "scrub_level": "standard"}

        plan_steps = plan.build_plan(inventory, options)
        categories = {step["category"] for step in plan_steps}
        assert "task-cleanup" in categories
        assert "filesystem-cleanup" in categories

    def test_scrub_level_aggressive_enables_deep_cleanup(self) -> None:
        """!
        @brief scrub_level='aggressive' enables shortcuts, COM, shell extensions.
        """
        inventory: dict[str, list[dict]] = {}
        options = {"auto_all": True, "scrub_level": "aggressive"}

        plan_steps = plan.build_plan(inventory, options)

        # Aggressive should auto-enable shortcuts
        fs_step = next((s for s in plan_steps if s["category"] == "filesystem-cleanup"), None)
        assert fs_step is not None
        assert fs_step["metadata"]["clean_shortcuts"] is True

        # Aggressive should auto-enable registry cleanup options
        reg_step = next((s for s in plan_steps if s["category"] == "registry-cleanup"), None)
        assert reg_step is not None
        assert reg_step["metadata"]["clean_addin_registry"] is True
        assert reg_step["metadata"]["clean_com_registry"] is True
        assert reg_step["metadata"]["clean_shell_extensions"] is True

        # Aggressive should add vnext-identity-cleanup
        assert any(s["category"] == "vnext-identity-cleanup" for s in plan_steps)

    def test_scrub_level_nuclear_enables_everything(self) -> None:
        """!
        @brief scrub_level='nuclear' enables ALL cleanup operations.
        """
        inventory: dict[str, list[dict]] = {}
        options = {"auto_all": True, "scrub_level": "nuclear"}

        plan_steps = plan.build_plan(inventory, options)
        categories = {step["category"] for step in plan_steps}

        # Nuclear should include advanced cleanup steps
        assert "vnext-identity-cleanup" in categories
        assert "taskband-cleanup" in categories
        assert "published-components-cleanup" in categories

        # Nuclear should enable all filesystem cleanup options
        fs_step = next((s for s in plan_steps if s["category"] == "filesystem-cleanup"), None)
        assert fs_step is not None
        assert fs_step["metadata"]["clean_msocache"] is True
        assert fs_step["metadata"]["clean_appx"] is True
        assert fs_step["metadata"]["clean_shortcuts"] is True

        # Nuclear should enable all registry cleanup options
        reg_step = next((s for s in plan_steps if s["category"] == "registry-cleanup"), None)
        assert reg_step is not None
        assert reg_step["metadata"]["clean_wi_metadata"] is True
        assert reg_step["metadata"]["remove_vba"] is True
        assert reg_step["metadata"]["clean_typelibs"] is True
        assert reg_step["metadata"]["clean_protocol_handlers"] is True

        # Nuclear should enable all license cleanup
        lic_step = next((s for s in plan_steps if s["category"] == "licensing-cleanup"), None)
        assert lic_step is not None
        assert lic_step["metadata"]["clean_spp"] is True
        assert lic_step["metadata"]["clean_ospp"] is True
        assert lic_step["metadata"]["clean_vnext"] is True
        assert lic_step["metadata"]["clean_all_licenses"] is True
