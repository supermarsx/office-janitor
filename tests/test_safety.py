"""!
@brief Safety guardrail enforcement tests.
@details Exercises dry-run propagation, whitelist enforcement, and targeted
scrub refusal policies applied by the safety module.
"""

from __future__ import annotations

import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import safety  # noqa: E402


@pytest.fixture(autouse=True)
def _mock_disk_usage_default(monkeypatch: pytest.MonkeyPatch) -> None:
    """!
    @brief Ensure runtime guards observe plentiful free disk space by default.
    @details Individual tests override the monkeypatch when simulating low space
    scenarios. The baseline keeps unrelated tests deterministic regardless of
    host disk utilisation.
    """

    def fake_usage(_: str) -> SimpleNamespace:
        baseline = safety.DEFAULT_MINIMUM_FREE_SPACE_BYTES * 4
        return SimpleNamespace(total=baseline * 2, used=baseline, free=baseline)

    monkeypatch.setattr(safety, "_query_disk_usage", fake_usage)


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
                    "paths": [r"C:\Program Files\Microsoft Office"],
                    "dry_run": False,
                },
            },
            {
                "id": "registry-0",
                "category": "registry-cleanup",
                "metadata": {
                    "keys": [r"HKLM\SOFTWARE\Microsoft\Office\16.0"],
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

    def test_preflight_blocks_template_cleanup_without_consent(self) -> None:
        """!
        @brief User template cleanup must be explicitly authorised.
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
                    "options": {"force": False, "keep_templates": False},
                },
            },
            {
                "id": "filesystem-0",
                "category": "filesystem-cleanup",
                "metadata": {
                    "paths": [r"C:\\Users\\Alice\\AppData\\Roaming\\Microsoft\\Templates"],
                    "dry_run": False,
                    "purge_templates": False,
                },
            },
        ]

        with pytest.raises(ValueError):
            safety.perform_preflight_checks(plan_steps)

    def test_preflight_allows_forced_template_cleanup(self) -> None:
        """!
        @brief Force flag should permit template cleanup steps.
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
                    "options": {"force": True, "keep_templates": False},
                },
            },
            {
                "id": "filesystem-0",
                "category": "filesystem-cleanup",
                "metadata": {
                    "paths": [r"C:\\Users\\Alice\\AppData\\Roaming\\Microsoft\\Templates"],
                    "dry_run": False,
                    "purge_templates": True,
                },
            },
        ]

        safety.perform_preflight_checks(plan_steps)

    def test_preflight_honours_preserve_templates_flag(self) -> None:
        """!
        @brief Preserve flag should block template deletion even with cleanup steps present.
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
                    "options": {"force": False, "keep_templates": True},
                },
            },
            {
                "id": "filesystem-0",
                "category": "filesystem-cleanup",
                "metadata": {
                    "paths": [r"C:\\Users\\Alice\\AppData\\Roaming\\Microsoft\\Templates"],
                    "dry_run": False,
                    "preserve_templates": True,
                },
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


class TestSafetyRuntimeEnvironment:
    """!
    @brief Runtime guard evaluation scenarios.
    @details Exercises administrative, OS, process, restore point, and dry-run
    enforcement helpers exposed by the safety module.
    """

    def test_runtime_guard_accepts_supported_environment(self) -> None:
        """!
        @brief Baseline acceptance for supported Windows releases.
        """

        safety.evaluate_runtime_environment(
            is_admin=True,
            os_system="Windows",
            os_release="10.0.19045",
            blocking_processes=[],
            dry_run=False,
            require_restore_point=True,
            restore_point_available=True,
        )

    def test_runtime_guard_requires_admin_when_not_dry_run(self) -> None:
        """!
        @brief Administrative rights are mandatory for destructive runs.
        """

        with pytest.raises(PermissionError):
            safety.evaluate_runtime_environment(
                is_admin=False,
                os_system="Windows",
                os_release="10.0",
                blocking_processes=[],
                dry_run=False,
                require_restore_point=False,
                restore_point_available=True,
            )

    def test_runtime_guard_allows_non_admin_dry_run(self) -> None:
        """!
        @brief Dry-run mode skips the administrative guard.
        """

        safety.evaluate_runtime_environment(
            is_admin=False,
            os_system="Windows",
            os_release="10.0",
            blocking_processes=[],
            dry_run=True,
            require_restore_point=False,
            restore_point_available=False,
        )

    def test_runtime_guard_rejects_unsupported_os(self) -> None:
        """!
        @brief Windows releases prior to 6.1 are blocked.
        """

        with pytest.raises(RuntimeError):
            safety.evaluate_runtime_environment(
                is_admin=True,
                os_system="Windows",
                os_release="5.1",
                blocking_processes=[],
                dry_run=False,
                require_restore_point=False,
                restore_point_available=True,
            )

    def test_runtime_guard_force_allows_unsupported_os(self) -> None:
        """!
        @brief Force flag bypasses the OS version guard.
        """

        safety.evaluate_runtime_environment(
            is_admin=True,
            os_system="Windows",
            os_release="5.1",
            blocking_processes=[],
            dry_run=False,
            require_restore_point=False,
            restore_point_available=True,
            force=True,
        )

    def test_runtime_guard_allow_flag_enables_unsupported_os(self) -> None:
        """!
        @brief Explicit override flag should bypass only the Windows guard.
        """

        safety.evaluate_runtime_environment(
            is_admin=True,
            os_system="Windows",
            os_release="5.1",
            blocking_processes=[],
            dry_run=False,
            require_restore_point=False,
            restore_point_available=True,
            allow_unsupported_windows=True,
        )

    def test_runtime_guard_blocks_lingering_processes(self) -> None:
        """!
        @brief Lingering Office processes prevent destructive actions.
        """

        with pytest.raises(RuntimeError):
            safety.evaluate_runtime_environment(
                is_admin=True,
                os_system="Windows",
                os_release="10.0",
                blocking_processes=["WINWORD.EXE"],
                dry_run=False,
                require_restore_point=False,
                restore_point_available=True,
            )

    def test_runtime_guard_force_allows_lingering_processes(self) -> None:
        """!
        @brief Force flag bypasses process blocks.
        """

        safety.evaluate_runtime_environment(
            is_admin=True,
            os_system="Windows",
            os_release="10.0",
            blocking_processes=["WINWORD.EXE"],
            dry_run=False,
            require_restore_point=False,
            restore_point_available=True,
            force=True,
        )

    def test_runtime_guard_dry_run_ignores_process_blocks(self) -> None:
        """!
        @brief Dry-run mode ignores process guard failures.
        """

        safety.evaluate_runtime_environment(
            is_admin=True,
            os_system="Windows",
            os_release="10.0",
            blocking_processes=["WINWORD.EXE"],
            dry_run=True,
            require_restore_point=False,
            restore_point_available=True,
        )

    def test_runtime_guard_rejects_non_windows_system(self) -> None:
        """!
        @brief Non-Windows operating systems should be rejected unless forced.
        """

        with pytest.raises(RuntimeError):
            safety.evaluate_runtime_environment(
                is_admin=True,
                os_system="Linux",
                os_release="5.15",
                blocking_processes=[],
                dry_run=False,
                require_restore_point=False,
                restore_point_available=True,
            )

    def test_runtime_guard_requires_restore_point(self) -> None:
        """!
        @brief Restore point requirement is enforced when enabled.
        """

        with pytest.raises(RuntimeError):
            safety.evaluate_runtime_environment(
                is_admin=True,
                os_system="Windows",
                os_release="10.0",
                blocking_processes=[],
                dry_run=False,
                require_restore_point=True,
                restore_point_available=False,
            )

    def test_runtime_guard_force_bypasses_restore_point(self) -> None:
        """!
        @brief Force flag allows proceeding without restore points.
        """

        safety.evaluate_runtime_environment(
            is_admin=True,
            os_system="Windows",
            os_release="10.0",
            blocking_processes=[],
            dry_run=False,
            require_restore_point=True,
            restore_point_available=False,
            force=True,
        )

    def test_runtime_guard_accepts_sufficient_free_space(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """!
        @brief Free-space guard succeeds when enough capacity remains.
        @details Monkeypatched disk usage is set to guarantee the guard observes
        adequate free space relative to the configured threshold override.
        """

        monkeypatch.setattr(
            safety,
            "_query_disk_usage",
            lambda _: SimpleNamespace(total=1000, used=100, free=900),
        )

        safety.evaluate_runtime_environment(
            is_admin=True,
            os_system="Windows",
            os_release="10.0",
            blocking_processes=[],
            dry_run=False,
            require_restore_point=False,
            restore_point_available=True,
            minimum_free_space_bytes=512,
        )

    def test_runtime_guard_rejects_insufficient_free_space(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """!
        @brief Guard raises when remaining free space is below the threshold.
        """

        monkeypatch.setattr(
            safety,
            "_query_disk_usage",
            lambda _: SimpleNamespace(total=1000, used=990, free=10),
        )

        with pytest.raises(RuntimeError):
            safety.evaluate_runtime_environment(
                is_admin=True,
                os_system="Windows",
                os_release="10.0",
                blocking_processes=[],
                dry_run=False,
                require_restore_point=False,
                restore_point_available=True,
                minimum_free_space_bytes=128,
            )

    def test_runtime_guard_force_bypasses_free_space(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """!
        @brief Force flag allows execution despite low free space.
        """

        monkeypatch.setattr(
            safety,
            "_query_disk_usage",
            lambda _: SimpleNamespace(total=1000, used=995, free=5),
        )

        safety.evaluate_runtime_environment(
            is_admin=True,
            os_system="Windows",
            os_release="10.0",
            blocking_processes=[],
            dry_run=False,
            require_restore_point=False,
            restore_point_available=True,
            minimum_free_space_bytes=256,
            force=True,
        )

    def test_guard_destructive_action_blocks_dry_run(self) -> None:
        """!
        @brief Dry-run prevents destructive operations.
        """

        with pytest.raises(RuntimeError):
            safety.guard_destructive_action(
                "delete files",
                dry_run=True,
            )

    def test_guard_destructive_action_respects_force(self) -> None:
        """!
        @brief Force flag overrides dry-run enforcement.
        """

        safety.guard_destructive_action(
            "delete files",
            dry_run=True,
            force=True,
        )

    def test_guard_destructive_action_allows_live_run(self) -> None:
        """!
        @brief Non dry-run mode allows destructive actions.
        """

        safety.guard_destructive_action(
            "delete files",
            dry_run=False,
        )

    def test_should_execute_destructive_action_returns_false_for_dry_run(self) -> None:
        """!
        @brief Boolean helper should block dry-run destructive actions.
        """

        assert (
            safety.should_execute_destructive_action(
                "delete files",
                dry_run=True,
            )
            is False
        )

    def test_should_execute_destructive_action_returns_true_for_live_run(self) -> None:
        """!
        @brief Boolean helper should allow non dry-run destructive actions.
        """

        assert (
            safety.should_execute_destructive_action(
                "delete files",
                dry_run=False,
            )
            is True
        )


class TestWhitelistBypass:
    """!
    @brief Tests for --no-whitelist / --dangerous-actions whitelist bypass.
    @details Ensures whitelist enforcement is correctly bypassed when both flags
    are active and that the bypass is rejected when only one flag is set.
    """

    def _make_plan_with_options(
        self,
        *,
        no_whitelist: bool = False,
        dangerous_actions: bool = False,
        registry_key: str = r"HKLM\SOFTWARE\Microsoft\SomethingCustom",
    ) -> list[dict[str, object]]:
        """Build a minimal plan with the given options and a non-whitelisted key."""
        return [
            {
                "id": "context",
                "category": "context",
                "metadata": {
                    "dry_run": False,
                    "mode": "auto-all",
                    "target_versions": [],
                    "unsupported_targets": [],
                    "options": {
                        "force": False,
                        "keep_templates": False,
                        "no_whitelist": no_whitelist,
                        "dangerous_actions": dangerous_actions,
                    },
                },
            },
            {
                "id": "registry-0",
                "category": "registry-cleanup",
                "metadata": {
                    "keys": [registry_key],
                    "dry_run": False,
                },
            },
        ]

    def test_preflight_blocks_non_whitelisted_key_by_default(self) -> None:
        """!
        @brief Non-whitelisted registry keys are rejected without bypass flags.
        """
        plan = self._make_plan_with_options()
        with pytest.raises(ValueError, match="non-whitelisted"):
            safety.perform_preflight_checks(plan)

    def test_preflight_blocks_no_whitelist_without_dangerous_actions(self) -> None:
        """!
        @brief --no-whitelist alone does not bypass the whitelist check.
        """
        plan = self._make_plan_with_options(no_whitelist=True, dangerous_actions=False)
        with pytest.raises(ValueError, match="non-whitelisted"):
            safety.perform_preflight_checks(plan)

    def test_preflight_blocks_dangerous_actions_without_no_whitelist(self) -> None:
        """!
        @brief --dangerous-actions alone does not bypass the whitelist check.
        """
        plan = self._make_plan_with_options(no_whitelist=False, dangerous_actions=True)
        with pytest.raises(ValueError, match="non-whitelisted"):
            safety.perform_preflight_checks(plan)

    def test_preflight_allows_bypass_with_both_flags(self) -> None:
        """!
        @brief Both --no-whitelist and --dangerous-actions bypass the whitelist.
        """
        plan = self._make_plan_with_options(no_whitelist=True, dangerous_actions=True)
        # Should NOT raise
        safety.perform_preflight_checks(plan)

    def test_preflight_bypass_still_enforces_dry_run_consistency(self) -> None:
        """!
        @brief Whitelist bypass does not disable other safety checks.
        """
        plan = self._make_plan_with_options(no_whitelist=True, dangerous_actions=True)
        # Introduce a dry-run mismatch
        plan[1]["metadata"]["dry_run"] = True  # type: ignore[index]
        with pytest.raises(ValueError, match="dry-run"):
            safety.perform_preflight_checks(plan)

    def test_preflight_bypass_logs_warning(self, caplog: pytest.LogCaptureFixture) -> None:
        """!
        @brief Bypassing the whitelist emits a warning about the dangerous mode.
        """
        import logging

        plan = self._make_plan_with_options(no_whitelist=True, dangerous_actions=True)
        with caplog.at_level(logging.WARNING):
            safety.perform_preflight_checks(plan)
        assert "DANGEROUS" in caplog.text

    def test_app_paths_are_now_whitelisted(self) -> None:
        """!
        @brief Office App Paths keys should be accepted by standard preflight.
        """
        plan = [
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
                "id": "registry-0",
                "category": "registry-cleanup",
                "metadata": {
                    "keys": [
                        r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion"
                        r"\App Paths\excel.exe",
                        r"HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows"
                        r"\CurrentVersion\App Paths\winword.exe",
                    ],
                    "dry_run": False,
                },
            },
        ]
        # Should NOT raise — these paths are now whitelisted
        safety.perform_preflight_checks(plan)

    def test_appv_keys_are_now_whitelisted(self) -> None:
        """!
        @brief App-V client registry keys should be accepted by standard preflight.
        """
        plan = [
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
                "id": "registry-0",
                "category": "registry-cleanup",
                "metadata": {
                    "keys": [
                        r"HKLM\SOFTWARE\Microsoft\AppV\Client\Packages",
                        r"HKLM\SOFTWARE\Microsoft\AppV\Client\Streaming\Packages",
                        r"HKLM\SOFTWARE\Microsoft\AppVISV",
                    ],
                    "dry_run": False,
                },
            },
        ]
        # Should NOT raise — these paths are now whitelisted
        safety.perform_preflight_checks(plan)
