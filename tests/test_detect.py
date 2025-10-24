"""!
@brief Detection scaffolding tests.
@details Coverage for registry probing and detection heuristics defined in the spec.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Dict, List

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import constants, detect


class TestRegistryDetectionScenarios:
    """!
    @brief Registry probing detection scenarios.
    @details Validates discovery of Office installations across registry hives, install roots, and
    release channels.
    """

    def test_msi_detection_aggregates_known_product_codes(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """!
        @brief Validate MSI discovery for multiple generations and architectures.
        """

        base_uninstall = constants.MSI_UNINSTALL_ROOTS[0][1]
        wow_uninstall = constants.MSI_UNINSTALL_ROOTS[1][1]

        def fake_iter_subkeys(root: int, path: str) -> List[str]:
            if path == base_uninstall:
                return [
                    "{91160000-0011-0000-0000-0000000FF1CE}",
                    "{91190000-0011-0000-0000-0000000FF1CE}",
                    "{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}",
                ]
            if path == wow_uninstall:
                return [
                    "{90150000-0011-0000-0000-0000000FF1CE}",
                    "{91140000-0011-0000-0000-0000000FF1CE}",
                ]
            return []

        key_values: Dict[str, Dict[str, str]] = {
            f"{base_uninstall}\\{{91160000-0011-0000-0000-0000000FF1CE}}": {
                "ProductCode": "{91160000-0011-0000-0000-0000000FF1CE}",
                "DisplayName": "Microsoft Office Professional Plus 2016",
                "DisplayVersion": "16.0.4266.1003",
            },
            f"{base_uninstall}\\{{91190000-0011-0000-0000-0000000FF1CE}}": {
                "ProductCode": "{91190000-0011-0000-0000-0000000FF1CE}",
                "DisplayName": "Microsoft Office Professional Plus 2019",
                "DisplayVersion": "16.0.10396.20017",
            },
            f"{base_uninstall}\\{{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}}": {
                "ProductCode": "{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}",
                "DisplayName": "Contoso Helper",
                "DisplayVersion": "1.0.0",
            },
            f"{wow_uninstall}\\{{90150000-0011-0000-0000-0000000FF1CE}}": {
                "ProductCode": "{90150000-0011-0000-0000-0000000FF1CE}",
                "DisplayName": "Microsoft Office Professional Plus 2013",
                "DisplayVersion": "15.0.4569.1506",
            },
            f"{wow_uninstall}\\{{91140000-0011-0000-0000-0000000FF1CE}}": {
                "ProductCode": "{91140000-0011-0000-0000-0000000FF1CE}",
                "DisplayName": "Microsoft Office Professional Plus 2010",
                "DisplayVersion": "14.0.7268.5000",
            },
        }

        def fake_read_values(root: int, path: str) -> Dict[str, str]:
            return key_values.get(path, {})

        monkeypatch.setattr(detect.registry_tools, "iter_subkeys", fake_iter_subkeys)
        monkeypatch.setattr(detect.registry_tools, "read_values", fake_read_values)

        installations = detect.detect_msi_installations()

        codes = {entry["product_code"] for entry in installations}
        assert codes == {
            "{91160000-0011-0000-0000-0000000FF1CE}",
            "{91190000-0011-0000-0000-0000000FF1CE}",
            "{90150000-0011-0000-0000-0000000FF1CE}",
            "{91140000-0011-0000-0000-0000000FF1CE}",
        }
        assert {entry["architecture"] for entry in installations} == {"x64", "x86"}
        assert all(entry["channel"] == "MSI" for entry in installations)

    def test_click_to_run_detection_collects_channel_metadata(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """!
        @brief Validate Click-to-Run discovery of channels, subscriptions, and COM registrations.
        """

        config_root, config_path = constants.C2R_CONFIGURATION_KEYS[0]
        subscription_path = constants.C2R_SUBSCRIPTION_ROOTS[0][1]
        com_path = constants.C2R_COM_REGISTRY_PATHS[0][1]

        def fake_read_values(root: int, path: str) -> Dict[str, str]:
            if path == config_path and root == config_root:
                return {
                    "ProductReleaseIds": "O365ProPlusRetail,ProjectProRetail",
                    "Platform": "x64",
                    "VersionToReport": "16.0.17029.20108",
                    "UpdateChannel": "http://officecdn.microsoft.com/pr/55336b82-a18d-4dd6-b5f6-9e5095c314a6",
                }
            if path == f"{subscription_path}\\O365ProPlusRetail":
                return {"ChannelId": "Production::MEC"}
            if path == f"{subscription_path}\\ProjectProRetail":
                return {"ChannelId": "Production::CC"}
            return {}

        def fake_iter_subkeys(root: int, path: str) -> List[str]:
            if path == subscription_path and root == constants.C2R_SUBSCRIPTION_ROOTS[0][0]:
                return ["O365ProPlusRetail", "ProjectProRetail"]
            if path == com_path and root == constants.C2R_COM_REGISTRY_PATHS[0][0]:
                return ["{1111}", "{2222}", "{3333}"]
            return []

        monkeypatch.setattr(detect.registry_tools, "read_values", fake_read_values)
        monkeypatch.setattr(detect.registry_tools, "iter_subkeys", fake_iter_subkeys)

        installations = detect.detect_c2r_installations()

        assert len(installations) == 1
        record = installations[0]
        assert record["release_ids"] == ["O365ProPlusRetail", "ProjectProRetail"]
        assert record["architecture"] == "x64"
        assert record["channel"] == "Monthly Enterprise Channel"
        assert record["com_registration_count"] == 3
        subscription_channels = {sub["channel"] for sub in record["subscriptions"]}
        assert subscription_channels == {"Monthly Enterprise Channel", "Current Channel"}

    def test_inventory_aggregates_registry_and_filesystem_signals(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """!
        @brief Validate that the inventory collector merges registry and filesystem hints.
        """

        monkeypatch.setattr(
            detect,
            "detect_msi_installations",
            lambda: [
                {
                    "product_code": "{91160000-0011-0000-0000-0000000FF1CE}",
                    "architecture": "x64",
                    "channel": "MSI",
                }
            ],
        )
        monkeypatch.setattr(
            detect,
            "detect_c2r_installations",
            lambda: [
                {
                    "release_ids": ["O365ProPlusRetail"],
                    "architecture": "x64",
                    "channel": "Current Channel",
                }
            ],
        )

        valid_paths = {
            constants.INSTALL_ROOT_TEMPLATES[0]["path"],
            constants.INSTALL_ROOT_TEMPLATES[2]["path"],
        }

        def fake_exists(self: Path) -> bool:  # type: ignore[override]
            return str(self) in valid_paths

        monkeypatch.setattr(detect.Path, "exists", fake_exists, raising=False)

        inventory = detect.gather_office_inventory()

        assert len(inventory["msi"]) == 1
        assert len(inventory["c2r"]) == 1
        assert len(inventory["filesystem"]) == 2
        labels = {entry["label"] for entry in inventory["filesystem"]}
        assert labels == {"c2r_root_x86", "office16_x86"}
