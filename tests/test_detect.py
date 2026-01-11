"""!
@brief Detection scaffolding tests.
@details Coverage for registry probing and detection heuristics defined in the spec.
"""

from __future__ import annotations

import json
import logging
import os
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import constants, detect, main, registry_tools  # noqa: E402


@pytest.fixture(autouse=True)
def stub_expensive_probes(monkeypatch: pytest.MonkeyPatch) -> None:
    """Prevent slow WMI/PowerShell probes from running during unit tests."""

    monkeypatch.setattr(detect, "_probe_msi_wmi", lambda: {})
    monkeypatch.setattr(detect, "_probe_msi_powershell", lambda: {})


@pytest.fixture
def msi_registry_layout(monkeypatch: pytest.MonkeyPatch) -> None:
    """!
    @brief Provide a fake registry layout for MSI product codes.
    """

    known_values: dict[tuple[int, str], dict[str, str]] = {}
    for product_code in (
        "{90160000-0011-0000-0000-0000000FF1CE}",
        "{90160000-0011-0000-1000-0000000FF1CE}",
        "{90150000-0011-0000-0000-0000000FF1CE}",
    ):
        for hive, base in constants.MSI_UNINSTALL_ROOTS:
            key = f"{base}\\{product_code}"
            known_values[(hive, key)] = {
                "ProductCode": product_code,
                "DisplayName": f"Display for {product_code}",
                "DisplayVersion": "16.0.0.123",
                "UninstallString": f"MsiExec.exe /X{product_code}",
                "DisplayIcon": (
                    r"C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE16"
                    r"\\Office Setup Controller\\setup.exe,0"
                ),
            }

    def fake_read_values(root: int, path: str) -> dict[str, str]:
        return known_values.get((root, path), {})

    monkeypatch.setattr(detect.registry_tools, "read_values", fake_read_values)


@pytest.fixture
def c2r_registry_layout(monkeypatch: pytest.MonkeyPatch) -> None:
    """!
    @brief Provide a fake registry layout for Click-to-Run metadata.
    """

    config_root, config_path = constants.C2R_CONFIGURATION_KEYS[0]
    subscription_root, subscription_path = constants.C2R_SUBSCRIPTION_ROOTS[0]
    release_root, release_path = constants.C2R_PRODUCT_RELEASE_ROOTS[0]

    known_values: dict[tuple[int, str], dict[str, str]] = {
        (config_root, config_path): {
            "ProductReleaseIds": "O365ProPlusRetail,ProjectProRetail",
            "Platform": "x64",
            "VersionToReport": "16.0.17029.20108",
            "UpdateChannel": "http://officecdn.microsoft.com/pr/55336b82-a18d-4dd6-b5f6-9e5095c314a6",
            "PackageGUID": "{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}",
            "InstallPath": r"C:\\Program Files\\Microsoft Office\\root",
        },
        (subscription_root, f"{subscription_path}\\O365ProPlusRetail"): {
            "ChannelId": "Production::MEC"
        },
        (subscription_root, f"{subscription_path}\\ProjectProRetail"): {
            "ChannelId": "Production::CC"
        },
        (release_root, f"{release_path}\\O365ProPlusRetail"): {},
        (release_root, f"{release_path}\\ProjectProRetail"): {},
    }

    def fake_read_values(root: int, path: str) -> dict[str, str]:
        return known_values.get((root, path), {})

    def fake_key_exists(root: int, path: str) -> bool:
        return (root, path) in known_values

    monkeypatch.setattr(detect.registry_tools, "read_values", fake_read_values)
    monkeypatch.setattr(detect.registry_tools, "key_exists", fake_key_exists)


class TestRegistryDetectionScenarios:
    """!
    @brief Registry probing detection scenarios.
    @details Validates discovery of Office installations across registry hives, install roots, and
    release channels.
    """

    def test_msi_detection_aggregates_known_product_codes(self, msi_registry_layout: None) -> None:
        """!
        @brief Validate MSI discovery for multiple generations and architectures.
        """

        installations = detect.detect_msi_installations()

        codes = {entry.product_code for entry in installations}
        expected_codes = {
            "{90160000-0011-0000-0000-0000000FF1CE}",
            "{90160000-0011-0000-1000-0000000FF1CE}",
            "{90150000-0011-0000-0000-0000000FF1CE}",
        }
        assert expected_codes.issubset(codes)
        assert all(entry.channel == "MSI" for entry in installations)
        assert {
            entry.architecture for entry in installations if entry.product_code in expected_codes
        } == {
            "x86",
            "x64",
        }

    def test_msi_detection_records_display_icon(self, msi_registry_layout: None) -> None:
        """!
        @brief Ensure MSI detection captures DisplayIcon and setup candidates.
        """

        installations = detect.detect_msi_installations()
        target_code = "{90160000-0011-0000-0000-0000000FF1CE}"
        record = next(entry for entry in installations if entry.product_code == target_code)
        expected_icon = (
            r"C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE16"
            r"\\Office Setup Controller\\setup.exe,0"
        )
        expected_setup = (
            r"C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE16"
            r"\\Office Setup Controller\\setup.exe"
        )
        assert record.display_icon == expected_icon
        assert record.maintenance_paths == (expected_setup,)
        assert record.properties["display_icon"] == expected_icon
        assert record.properties["maintenance_paths"] == [expected_setup]

    def test_click_to_run_detection_collects_channel_metadata(
        self, c2r_registry_layout: None
    ) -> None:
        """!
        @brief Validate Click-to-Run discovery of channels, subscriptions, and COM registrations.
        """

        installations = detect.detect_c2r_installations()

        matching = [entry for entry in installations if entry.release_ids == ("O365ProPlusRetail",)]
        assert matching, "Expected to find O365ProPlusRetail release"
        record = matching[0]
        assert record.architecture == "x64"
        assert record.channel == "Monthly Enterprise Channel"
        assert record.properties["version"] == "16.0.17029.20108"
        assert record.properties["package_guid"] == "{AAAAAAAA-BBBB-CCCC-DDDD-EEEEEEEEEEEE}"
        assert record.properties["supported_architectures"]

    def test_inventory_aggregates_registry_and_filesystem_signals(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """!
        @brief Validate that the inventory collector merges registry and filesystem hints.
        """

        sample_msi = detect.DetectedInstallation(
            source="MSI",
            product="Microsoft Office Professional Plus 2016",
            version="2016",
            architecture="x64",
            uninstall_handles=(
                "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{90160000-0011-0000-1000-0000000FF1CE}",
            ),
            channel="MSI",
            product_code="{90160000-0011-0000-1000-0000000FF1CE}",
            properties={"display_name": "ProPlus", "display_version": "16.0.10396.20017"},
        )
        sample_c2r = detect.DetectedInstallation(
            source="C2R",
            product="Microsoft 365 Apps for enterprise",
            version="16.0.17029.20108",
            architecture="x64",
            uninstall_handles=("HKLM\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration",),
            channel="Current Channel",
            release_ids=("O365ProPlusRetail",),
            properties={"supported_versions": ["2016", "2019"], "supported_architectures": ["x64"]},
        )

        monkeypatch.setattr(detect, "detect_msi_installations", lambda: [sample_msi])
        monkeypatch.setattr(detect, "detect_c2r_installations", lambda: [sample_c2r])

        monkeypatch.setenv("PROGRAMDATA", r"C:\\ProgramData")
        monkeypatch.setenv("LOCALAPPDATA", r"C:\\Users\\Default\\AppData\\Local")
        monkeypatch.setenv("APPDATA", r"C:\\Users\\Default\\AppData\\Roaming")

        valid_paths = {
            str(Path(constants.INSTALL_ROOT_TEMPLATES[0]["path"])),
            str(Path(constants.INSTALL_ROOT_TEMPLATES[2]["path"])),
            *(
                str(Path(os.path.expandvars(template["path"])))
                for template in constants.RESIDUE_PATH_TEMPLATES
            ),
        }

        def fake_exists(self: Path) -> bool:  # type: ignore[override]
            return str(self) in valid_paths

        monkeypatch.setattr(detect.Path, "exists", fake_exists, raising=False)

        monkeypatch.setattr(
            detect,
            "gather_running_office_processes",
            lambda: [{"name": "winword.exe", "pid": "1234"}],
        )
        monkeypatch.setattr(
            detect,
            "gather_office_services",
            lambda: [{"name": "ClickToRunSvc", "state": "RUNNING"}],
        )
        monkeypatch.setattr(
            detect,
            "gather_office_tasks",
            lambda: [
                {
                    "task": r"\Microsoft\Office\OfficeTelemetryAgentLogOn",
                    "status": "Ready",
                }
            ],
        )
        monkeypatch.setattr(
            detect,
            "gather_activation_state",
            lambda: {
                "path": constants.OSPP_REGISTRY_PATH,
                "values": {"SKUID": "Test"},
            },
        )
        monkeypatch.setattr(
            detect,
            "gather_registry_residue",
            lambda: [{"path": r"HKLM\SOFTWARE\Microsoft\Office"}],
        )

        inventory = detect.gather_office_inventory()

        assert len(inventory["msi"]) == 1
        assert inventory["msi"][0]["product_code"] == "{90160000-0011-0000-1000-0000000FF1CE}"
        assert inventory["msi"][0]["properties"]["display_name"] == "ProPlus"
        assert inventory["msi"][0]["uninstall_handles"] == [
            "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{90160000-0011-0000-1000-0000000FF1CE}"
        ]
        assert len(inventory["c2r"]) == 1
        assert inventory["c2r"][0]["release_ids"] == ["O365ProPlusRetail"]
        assert inventory["c2r"][0]["properties"]["supported_architectures"] == ["x64"]
        assert len(inventory["filesystem"]) == len(valid_paths)
        labels = {entry["label"] for entry in inventory["filesystem"]}
        expected_labels = {
            "c2r_root_x86",
            "office16_x86",
            *[template["label"] for template in constants.RESIDUE_PATH_TEMPLATES],
        }
        assert labels == expected_labels
        assert {entry.get("architecture", "x86") for entry in inventory["filesystem"]} >= {"x86"}
        assert inventory["processes"] == [{"name": "winword.exe", "pid": "1234"}]
        assert inventory["services"] == [{"name": "ClickToRunSvc", "state": "RUNNING"}]
        assert inventory["tasks"][0]["task"] == r"\Microsoft\Office\OfficeTelemetryAgentLogOn"
        assert inventory["activation"]["path"] == constants.OSPP_REGISTRY_PATH
        assert inventory["registry"] == [{"path": r"HKLM\SOFTWARE\Microsoft\Office"}]

    def test_run_detection_persists_inventory_snapshot(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """!
        @brief Ensure detection snapshots are written to the resolved log directory.
        """

        sample_inventory = {
            "msi": [],
            "c2r": [],
            "filesystem": [],
            "processes": [],
            "services": [],
            "tasks": [],
            "activation": {},
            "registry": [],
        }

        monkeypatch.setattr(main.detect, "gather_office_inventory", lambda: sample_inventory)

        machine_log = logging.getLogger("office-janitor-test")
        machine_log.handlers.clear()
        machine_log.addHandler(logging.NullHandler())
        machine_log.propagate = False

        result = main._run_detection(machine_log, tmp_path)

        assert result is sample_inventory

        snapshots = list(tmp_path.glob("inventory-*.json"))
        assert len(snapshots) == 1
        payload = json.loads(snapshots[0].read_text(encoding="utf-8"))
        assert payload == sample_inventory

    def test_registry_residue_templates_cover_office_versions(
        self, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """!
        @brief Validate registry residue probing includes versioned Office hives.
        """

        present = {
            (constants.HKLM, r"SOFTWARE\Microsoft\Office\16.0"),
            (constants.HKLM, r"SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform"),
            (
                constants.HKLM,
                r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform"
                r"\0ff1ce15-a989-479d-af46-f275c6370663",
            ),
            (
                constants.HKLM,
                r"SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663",
            ),
        }

        monkeypatch.setattr(
            detect,
            "_key_exists_with_fallback",
            lambda hive, path: (hive, path) in present,
        )

        entries = detect.gather_registry_residue()
        handles = {entry["path"] for entry in entries}

        expected_prefix = registry_tools.hive_name(constants.HKLM)
        assert f"{expected_prefix}\\SOFTWARE\\Microsoft\\Office\\16.0" in handles
        assert any("SoftwareProtectionPlatform" in handle for handle in handles)


def test_reprobe_with_options(monkeypatch: pytest.MonkeyPatch) -> None:
    """!
    @brief Test reprobe accepts options mapping.
    """

    monkeypatch.setattr(detect, "_probe_msi_wmi", lambda: {})
    monkeypatch.setattr(detect, "_probe_msi_powershell", lambda: {})

    options = {"limited_user": True}
    result = detect.reprobe(options)

    assert isinstance(result, dict)
    assert "msi" in result
    assert "c2r" in result


def test_detect_msi_installations_with_registry(monkeypatch: pytest.MonkeyPatch) -> None:
    """!
    @brief Test MSI detection with fake registry entries.
    """

    monkeypatch.setattr(detect, "_probe_msi_wmi", lambda: {})
    monkeypatch.setattr(detect, "_probe_msi_powershell", lambda: {})

    # Mock registry values
    fake_values = {
        "ProductCode": "{90160000-0011-0000-0000-0000000FF1CE}",
        "DisplayName": "Microsoft Office Professional Plus 2016",
        "DisplayVersion": "16.0.1234.5678",
        "UninstallString": 'MsiExec.exe /X{90160000-0011-0000-0000-0000000FF1CE}',
    }

    monkeypatch.setattr(
        detect,
        "_read_values_with_fallback",
        lambda hive, path: fake_values if "Uninstall" in path and "{90160000" in path else {},
    )

    installations = detect.detect_msi_installations()

    assert len(installations) > 0
    found = next((inst for inst in installations if inst.product_code == "{90160000-0011-0000-0000-0000000FF1CE}"), None)
    assert found is not None
    assert found.product == "Microsoft Office Professional Plus 2016"
    assert found.version == "16.0.1234.5678"


def test_detect_c2r_installations_with_registry(monkeypatch: pytest.MonkeyPatch) -> None:
    """!
    @brief Test C2R detection with fake registry entries.
    """

    fake_config_values = {
        "ProductReleaseIds": "O365ProPlusRetail",
        "Platform": "x64",
        "VersionToReport": "16.0.12345.67890",
        "UpdateChannel": "http://officecdn.microsoft.com/pr/7ffbc6bf-bc32-4f92-8982-f9dd17fd3114",
    }

    monkeypatch.setattr(
        detect,
        "_read_values_with_fallback",
        lambda hive, path: fake_config_values if "Configuration" in path else {},
    )

    installations = detect.detect_c2r_installations()

    assert len(installations) > 0
    found = next((inst for inst in installations if "O365ProPlusRetail" in inst.release_ids), None)
    assert found is not None
    assert found.source == "C2R"
    assert found.version == "16.0.12345.67890"
