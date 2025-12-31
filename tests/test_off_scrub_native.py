from __future__ import annotations

from pathlib import Path
import logging

from office_janitor import off_scrub_native, logging_ext, tasks_services
import pytest


def test_parse_legacy_arguments_msi_flags(tmp_path):
    invocation = off_scrub_native._parse_legacy_arguments(
        "msi",
        [
            str(Path("C:/temp/OffScrub10.vbs")),
            "ALL",
            "/NOREBOOT",
            "/L",
            str(tmp_path / "logs"),
            "{90140000-0011-0000-0000-0000000FF1CE}",
        ],
    )

    assert invocation.version_group == "2010"
    assert invocation.flags.get("all") is True
    assert invocation.flags.get("no_reboot") is True
    assert invocation.log_directory == tmp_path / "logs"
    assert "{90140000-0011-0000-0000-0000000FF1CE}" in invocation.product_codes


def test_select_msi_targets_filters_by_group():
    inventory = {
        "msi": [
            {
                "product_code": "{AAA11111-2222-3333-4444-555555555555}",
                "version": "14.0.1234",
                "properties": {"supported_versions": ["2010"]},
            },
            {
                "product_code": "{BBB11111-2222-3333-4444-555555555555}",
                "version": "16.0.9999",
                "properties": {"supported_versions": ["2016"]},
            },
        ]
    }
    invocation = off_scrub_native.LegacyInvocation(
        script_path=None,
        version_group="2010",
        product_codes=[],
        release_ids=[],
        flags={"all": True},
        unknown=[],
    )

    targets = off_scrub_native._select_msi_targets(invocation, inventory)
    assert len(targets) == 1
    assert targets[0]["product_code"] == "{AAA11111-2222-3333-4444-555555555555}"


def test_select_c2r_targets_respects_release_ids():
    inventory = {
        "c2r": [
            {
                "release_ids": ["O365ProPlusRetail"],
                "product": "Microsoft 365 Apps for enterprise",
                "version": "16.0",
            }
        ]
    }
    invocation = off_scrub_native.LegacyInvocation(
        script_path=None,
        version_group="c2r",
        product_codes=[],
        release_ids=["O365ProPlusRetail"],
        flags={"all": False},
        unknown=[],
    )

    targets = off_scrub_native._select_c2r_targets(invocation, inventory)
    assert len(targets) == 1
    assert targets[0]["release_ids"] == ["O365ProPlusRetail"]


def test_test_rerun_runs_twice_for_msi(monkeypatch):
    inventory = {
        "msi": [
            {
                "product_code": "{AAA11111-2222-3333-4444-555555555555}",
                "version": "14.0.1234",
                "properties": {"supported_versions": ["2010"]},
            }
        ]
    }
    monkeypatch.setattr(off_scrub_native.detect, "gather_office_inventory", lambda: inventory)

    calls = []

    def fake_uninstall(products, dry_run=False, retries=None):
        calls.append(list(products))

    monkeypatch.setattr(off_scrub_native, "uninstall_msi_products", fake_uninstall)

    rc = off_scrub_native.main(["msi", "OffScrub10.vbs", "/TR", "ALL"])
    assert rc == 0
    assert len(calls) == 2


def test_offline_flag_carried_into_c2r_invocation(monkeypatch):
    inventory = {
        "c2r": [
            {"release_ids": ["O365ProPlusRetail"], "product": "Microsoft 365 Apps", "version": "16.0"}
        ]
    }
    monkeypatch.setattr(off_scrub_native.detect, "gather_office_inventory", lambda: inventory)

    captured = []

    def fake_uninstall(config, dry_run=False, retries=None):
        captured.append(config)

    monkeypatch.setattr(off_scrub_native, "uninstall_products", fake_uninstall)

    rc = off_scrub_native.main(["c2r", "OffScrubC2R.vbs", "/OFFLINE", "ALL"])
    assert rc == 0
    assert captured
    assert captured[0].get("offline") is True


def test_quiet_suppresses_info_logging(monkeypatch, caplog):
    inventory = {
        "c2r": [
            {"release_ids": ["O365ProPlusRetail"], "product": "Microsoft 365 Apps", "version": "16.0"}
        ]
    }
    monkeypatch.setattr(off_scrub_native.detect, "gather_office_inventory", lambda: inventory)

    monkeypatch.setattr(off_scrub_native, "uninstall_products", lambda config, dry_run=False, retries=None: None)

    caplog.set_level(logging.INFO, logger=logging_ext.HUMAN_LOGGER_NAME)
    rc = off_scrub_native.main(["c2r", "OffScrubC2R.vbs", "/QUIET", "ALL"])
    assert rc == 0
    assert not [record for record in caplog.records if record.levelno == logging.INFO]


def test_no_reboot_suppresses_recommendations(monkeypatch):
    inventory = {
        "c2r": [
            {"release_ids": ["O365ProPlusRetail"], "product": "Microsoft 365 Apps", "version": "16.0"}
        ]
    }
    monkeypatch.setattr(off_scrub_native.detect, "gather_office_inventory", lambda: inventory)

    tasks_services.consume_reboot_recommendations()

    def fake_uninstall(config, dry_run=False, retries=None):
        tasks_services._record_reboot_recommendation("ClickToRunSvc")  # type: ignore[attr-defined]

    monkeypatch.setattr(off_scrub_native, "uninstall_products", fake_uninstall)

    rc = off_scrub_native.main(["c2r", "OffScrubC2R.vbs", "/NOREBOOT", "ALL"])
    assert rc == 0
    assert tasks_services.consume_reboot_recommendations() == []


def test_user_settings_flags_forwarded(monkeypatch):
    inventory = {
        "msi": [
            {
                "product_code": "{AAA11111-2222-3333-4444-555555555555}",
                "version": "14.0.1234",
                "properties": {"supported_versions": ["2010"]},
            }
        ]
    }
    monkeypatch.setattr(off_scrub_native.detect, "gather_office_inventory", lambda: inventory)

    captured = []

    def fake_uninstall(products, dry_run=False, retries=None):
        captured.append(products)

    monkeypatch.setattr(off_scrub_native, "uninstall_msi_products", fake_uninstall)

    rc = off_scrub_native.main(["msi", "OffScrub10.vbs", "/DELETEUSERSETTINGS", "ALL"])
    assert rc == 0
    assert captured
    assert captured[0][0].get("delete_user_settings") is True


def test_user_settings_cleanup_executed(monkeypatch):
    inventory = {
        "msi": [
            {
                "product_code": "{AAA11111-2222-3333-4444-555555555555}",
                "version": "14.0.1234",
                "properties": {"supported_versions": ["2010"]},
            }
        ]
    }
    monkeypatch.setattr(off_scrub_native.detect, "gather_office_inventory", lambda: inventory)
    monkeypatch.setattr(off_scrub_native, "uninstall_msi_products", lambda products, dry_run=False, retries=None: None)

    removed = []

    def fake_remove_paths(paths, dry_run=False):
        removed.extend(paths)

    monkeypatch.setattr(off_scrub_native.fs_tools, "remove_paths", fake_remove_paths)

    rc = off_scrub_native.main(["msi", "OffScrub10.vbs", "/DELETEUSERSETTINGS", "ALL"])
    assert rc == 0
    assert removed


def test_clear_addin_registry_calls_delete(monkeypatch):
    inventory = {
        "c2r": [
            {"release_ids": ["O365ProPlusRetail"], "product": "Microsoft 365 Apps", "version": "16.0"}
        ]
    }
    monkeypatch.setattr(off_scrub_native.detect, "gather_office_inventory", lambda: inventory)
    monkeypatch.setattr(off_scrub_native, "uninstall_products", lambda config, dry_run=False, retries=None: None)

    deleted = []

    def fake_delete_keys(keys, dry_run=False, logger=None):
        deleted.extend(keys)

    monkeypatch.setattr(off_scrub_native.registry_tools, "delete_keys", fake_delete_keys)

    rc = off_scrub_native.main(["c2r", "OffScrubC2R.vbs", "/CLEARADDINREG", "ALL"])
    assert rc == 0
    assert deleted


def test_unmapped_flags_logged(monkeypatch, caplog):
    caplog.set_level("INFO")
    inventory = {
        "msi": [
            {
                "product_code": "{AAA11111-2222-3333-4444-555555555555}",
                "version": "14.0.1234",
                "properties": {"supported_versions": ["2010"]},
            }
        ]
    }
    monkeypatch.setattr(off_scrub_native.detect, "gather_office_inventory", lambda: inventory)
    monkeypatch.setattr(off_scrub_native, "uninstall_msi_products", lambda products, dry_run=False, retries=None: None)

    rc = off_scrub_native.main(["msi", "OffScrub10.vbs", "/NOREBOOT", "ALL"])
    assert rc == 0
    assert not any("not yet implemented" in record.message for record in caplog.records)
