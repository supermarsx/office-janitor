from __future__ import annotations

import sys
from pathlib import Path

import pytest

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import constants


class TestConstantsModule:
    """!
    @brief Validate helpers exposed from :mod:`office_janitor.constants`.
    """

    @pytest.mark.parametrize(
        "product_code,expected",
        [
            ("{90160000-0011-0000-0000-0000000FF1CE}", "office"),
            ("901600000011000000000000000FF1CE", "office"),
            ("{90160000-003B-0000-0000-0000000FF1CE}", "project"),
        ],
    )
    def test_resolve_msi_family(self, product_code: str, expected: str) -> None:
        """!
        @brief MSI product codes map to stable families.
        """

        assert constants.resolve_msi_family(product_code) == expected

    def test_resolve_c2r_family(self) -> None:
        """!
        @brief Click-to-Run releases expose family metadata.
        """

        assert constants.resolve_c2r_family("ProjectProRetail") == "project"
        assert constants.resolve_c2r_family("VisioProRetail") == "visio"

    @pytest.mark.parametrize(
        "component,expected",
        [
            ("visio", "visio"),
            ("MSI-Project", "project"),
            ("OneNote2016", "onenote"),
        ],
    )
    def test_component_normalisation(self, component: str, expected: str) -> None:
        """!
        @brief Optional component aliases resolve to supported identifiers.
        """

        assert constants.resolve_supported_component(component) == expected
        assert constants.is_supported_component(component) is True

    def test_supported_components_iterable(self) -> None:
        """!
        @brief Helper exposes default optional components.
        """

        entries = constants.iter_supported_components()
        assert set(entries) == set(constants.SUPPORTED_COMPONENTS)

    def test_uninstall_command_templates_expose_expected_metadata(self) -> None:
        """!
        @brief Uninstall command templates provide OffScrub wiring.
        """

        msi_template = constants.UNINSTALL_COMMAND_TEMPLATES["msi"]
        assert msi_template["executable"] == constants.OFFSCRUB_EXECUTABLE
        assert constants.MSI_OFFSCRUB_DEFAULT_SCRIPT in set(
            msi_template["script_map"].values()
        )

        c2r_template = constants.UNINSTALL_COMMAND_TEMPLATES["c2r"]
        assert c2r_template["script"] == constants.C2R_OFFSCRUB_SCRIPT
        assert tuple(c2r_template["arguments"]) == constants.C2R_OFFSCRUB_ARGS
