"""!
@brief Tests for ``scripts/release_pypi.py`` automation helpers.
"""

from __future__ import annotations

import pathlib
import subprocess
import sys

PROJECT_ROOT = pathlib.Path(__file__).resolve().parents[1]
SCRIPTS_PATH = PROJECT_ROOT / "scripts"
if str(SCRIPTS_PATH) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_PATH))

import release_pypi  # noqa: E402


def test_parse_patch_version_accepts_valid_value() -> None:
    """!
    @brief ``parse_patch_version`` should parse valid ``0.0.x`` strings.
    """

    assert release_pypi.parse_patch_version("0.0.17") == 17


def test_parse_patch_version_rejects_invalid_value() -> None:
    """!
    @brief ``parse_patch_version`` should reject non ``0.0.x`` formats.
    """

    try:
        release_pypi.parse_patch_version("26.17")
    except ValueError as exc:
        assert "Expected 0.0.x version" in str(exc)
    else:  # pragma: no cover - defensive
        raise AssertionError("Expected ValueError for non 0.0.x version")


def test_extract_published_patches_filters_other_versions() -> None:
    """!
    @brief PyPI metadata parser should keep only ``0.0.x`` patch entries.
    """

    payload = {
        "releases": {
            "0.0.0": [{}],
            "0.0.8": [{}],
            "0.1.0": [{}],
            "26.17": [{}],
            "0.0.8rc1": [{}],
        }
    }

    assert release_pypi.extract_published_patches(payload) == {0, 8}


def test_choose_next_patch_uses_highest_local_or_pypi() -> None:
    """!
    @brief Next patch should advance from the maximum local/published value.
    """

    assert release_pypi.choose_next_patch("0.0.2", {0, 1, 2, 3}) == 4
    assert release_pypi.choose_next_patch("0.0.9", {0, 1, 2, 3}) == 10


def test_is_existing_file_upload_error_matches_twine_collision_text() -> None:
    """!
    @brief Collision detector should classify existing-file upload failures.
    """

    exc = subprocess.CalledProcessError(
        returncode=1,
        cmd=["python", "-m", "twine", "upload"],
        output=(
            "HTTPError: 400 Bad Request from https://upload.pypi.org/legacy/ "
            "This filename has already been used"
        ),
        stderr="",
    )

    assert release_pypi.is_existing_file_upload_error(exc) is True
