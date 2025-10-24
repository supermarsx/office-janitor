"""!
@brief Safety guardrail enforcement tests.
@details Exercises dry-run propagation, whitelist enforcement, and targeted
scrub refusal policies applied by the safety module.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import safety


class TestSafetyPreflight:
    """!
    @brief Preflight safety validation scenarios.
    @details Ensures destructive steps stay behind guardrails and the planner's
    metadata is honored before execution proceeds.
    """

    def test_preflight_accepts_well_formed_plan(self) -> None:
        """!
        @brief Baseline validation for a compliant plan.
        @details No exception should be raised when every step resides within
        allowed paths, target versions, and dry-run semantics.
        """

        plan_steps = [
            {
                "id": "context",
                "category": "context",
                "metadata": {
                    "dry_run": False,
                    "mode": "auto-all",
                    "target_versions": ["2019"],
                    "unsupported_targets": [],
                },
            },
            {
                "id": "msi-0",
                "category": "msi-uninstall",
                "metadata": {"version": "2019", "dry_run": False},
            },
            {
                "id": "filesystem-0",
                "category": "filesystem-cleanup",
                "metadata": {
                    "paths": [r"C:\\Program Files\\Microsoft Office"],
                    "dry_run": False,
                },
            },
            {
                "id": "registry-0",
                "category": "registry-cleanup",
                "metadata": {
                    "keys": [r"HKLM\\SOFTWARE\\Microsoft\\Office\\16.0"],
                    "dry_run": False,
                },
            },
            {
                "id": "licensing-0",
                "category": "licensing-cleanup",
                "metadata": {"dry_run": False},
            },
        ]

        safety.perform_preflight_checks(plan_steps)

    def test_preflight_rejects_unsupported_target(self) -> None:
        """!
        @brief Unsupported versions trigger a refusal.
        @details Safety should prevent execution when context metadata cites
        versions outside the supported range.
        """

        plan_steps = [
            {
                "id": "context",
                "category": "context",
                "metadata": {
                    "dry_run": False,
                    "mode": "target:1999",
                    "target_versions": [],
                    "unsupported_targets": ["1999"],
                },
            }
        ]

        with pytest.raises(ValueError):
            safety.perform_preflight_checks(plan_steps)

    def test_preflight_rejects_mismatched_target_scope(self) -> None:
        """!
        @brief Targeted scrubs must only uninstall matching versions.
        @details A mismatch between uninstall step metadata and selected targets
        should be rejected.
        """

        plan_steps = [
            {
                "id": "context",
                "category": "context",
                "metadata": {
                    "dry_run": False,
                    "mode": "target:2019",
                    "target_versions": ["2019"],
                    "unsupported_targets": [],
                },
            },
            {
                "id": "msi-0",
                "category": "msi-uninstall",
                "metadata": {"version": "2016", "dry_run": False},
            },
        ]

        with pytest.raises(ValueError):
            safety.perform_preflight_checks(plan_steps)

    def test_preflight_rejects_blacklisted_filesystem(self) -> None:
        """!
        @brief Filesystem whitelist is enforced.
        @details Paths rooted in Windows system directories should trigger an
        immediate refusal.
        """

        plan_steps = [
            {
                "id": "context",
                "category": "context",
                "metadata": {
                    "dry_run": False,
                    "mode": "auto-all",
                    "target_versions": [],
                    "unsupported_targets": [],
                },
            },
            {
                "id": "filesystem-0",
                "category": "filesystem-cleanup",
                "metadata": {
                    "paths": [r"C:\\Windows\\Temp\\Office"],
                    "dry_run": False,
                },
            },
        ]

        with pytest.raises(ValueError):
            safety.perform_preflight_checks(plan_steps)

    def test_preflight_allows_whitelisted_user_profile_path(self) -> None:
        """!
        @brief User profile Office paths remain allowed despite broad blacklist.
        @details Regression coverage ensuring `%APPDATA%` expansions under
        `C:\\Users` are accepted even though the user directory is generally
        blocked.
        """

        plan_steps = [
            {
                "id": "context",
                "category": "context",
                "metadata": {
                    "dry_run": False,
                    "mode": "cleanup-only",
                    "target_versions": [],
                    "unsupported_targets": [],
                },
            },
            {
                "id": "filesystem-0",
                "category": "filesystem-cleanup",
                "metadata": {
                    "paths": [r"C:\\Users\\Alice\\AppData\\Roaming\\Microsoft\\Office"],
                    "dry_run": False,
                },
            },
        ]

        safety.perform_preflight_checks(plan_steps)

    def test_preflight_detects_dry_run_mismatch(self) -> None:
        """!
        @brief Dry-run flag must propagate consistently.
        @details If any step disagrees with the global dry-run selection the plan
        should be rejected.
        """

        plan_steps = [
            {
                "id": "context",
                "category": "context",
                "metadata": {
                    "dry_run": True,
                    "mode": "auto-all",
                    "target_versions": [],
                    "unsupported_targets": [],
                },
            },
            {
                "id": "licensing-0",
                "category": "licensing-cleanup",
                "metadata": {"dry_run": False},
            },
        ]

        with pytest.raises(ValueError):
            safety.perform_preflight_checks(plan_steps)

    def test_preflight_requires_targeted_uninstall(self) -> None:
        """!
        @brief Target mode without uninstall steps is invalid.
        @details Ensures targeted plans still reference at least one uninstall
        step before proceeding.
        """

        plan_steps = [
            {
                "id": "context",
                "category": "context",
                "metadata": {
                    "dry_run": False,
                    "mode": "target:2019",
                    "target_versions": ["2019"],
                    "unsupported_targets": [],
                },
            },
            {
                "id": "licensing-0",
                "category": "licensing-cleanup",
                "metadata": {"dry_run": False},
            },
        ]

        with pytest.raises(ValueError):
            safety.perform_preflight_checks(plan_steps)

