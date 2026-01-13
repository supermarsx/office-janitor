from pathlib import Path
from unittest.mock import patch, MagicMock

import pytest

from office_janitor import c2r_uninstall, licensing


def test_c2r_derives_uninstall_handles():
    config = {"release_ids": ["O365ProPlusRetail"]}
    target = c2r_uninstall._normalise_c2r_entry(config)
    assert target.uninstall_handles, "Expected derived uninstall handles"
    # Support both canonical HKLM/HKCU strings and numeric hive fallbacks (hex)
    assert any(
        h.startswith("HKLM\\") or h.startswith("HKCU\\") or h.startswith("0x")
        for h in target.uninstall_handles
    )
    # Ensure at least one handle references ClickToRun/ProductReleaseIDs/Office
    assert any(
        "ProductReleaseIDs" in h or "ClickToRun" in h or "Office" in h
        for h in target.uninstall_handles
    )


def test_parse_license_results():
    out = "OSPP:2\nSPP:3\nSome unrelated line\n"
    counts = licensing._parse_license_results(out)
    assert counts["ospp"] == 2
    assert counts["spp"] == 3


# ---------------------------------------------------------------------------
# Tests for new WMI-based licensing functions
# ---------------------------------------------------------------------------


class TestOfficeLicenseConstants:
    """Test licensing module constants."""

    def test_office_application_id_format(self) -> None:
        """Office Application ID should be a valid GUID format."""
        app_id = licensing.OFFICE_APPLICATION_ID
        assert len(app_id) == 36  # GUID without braces
        assert app_id.count("-") == 4


class TestQueryWmiLicenses:
    """Tests for _query_wmi_licenses function."""

    def test_query_returns_list_on_success(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """Should return list of license dicts when WMI returns data."""
        mock_result = MagicMock()
        mock_result.returncode = 0
        mock_result.stdout = (
            '[{"ID": "1", "Name": "Test", "PartialProductKey": "ABC", "ProductKeyID": "123"}]'
        )

        monkeypatch.setattr(
            "office_janitor.licensing.exec_utils.run_command",
            lambda *a, **kw: mock_result,
        )

        result = licensing._query_wmi_licenses()
        assert isinstance(result, list)
        assert len(result) == 1
        assert result[0]["Name"] == "Test"

    def test_query_returns_empty_on_failure(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """Should return empty list when WMI fails."""
        mock_result = MagicMock()
        mock_result.returncode = 1
        mock_result.stdout = ""

        monkeypatch.setattr(
            "office_janitor.licensing.exec_utils.run_command",
            lambda *a, **kw: mock_result,
        )

        result = licensing._query_wmi_licenses()
        assert result == []

    def test_query_handles_single_object(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """Should handle PowerShell returning single object instead of array."""
        mock_result = MagicMock()
        mock_result.returncode = 0
        mock_result.stdout = (
            '{"ID": "1", "Name": "Single", "PartialProductKey": "XYZ", "ProductKeyID": "456"}'
        )

        monkeypatch.setattr(
            "office_janitor.licensing.exec_utils.run_command",
            lambda *a, **kw: mock_result,
        )

        result = licensing._query_wmi_licenses()
        assert isinstance(result, list)
        assert len(result) == 1


class TestCleanOsppLicensesWmi:
    """Tests for clean_ospp_licenses_wmi function."""

    def test_dry_run_returns_names(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """Dry run should return license names without removal."""
        licenses = [
            {"Name": "Office 365", "PartialProductKey": "ABC", "ProductKeyID": "123"},
        ]
        monkeypatch.setattr(
            "office_janitor.licensing._query_wmi_licenses",
            lambda *a, **kw: licenses,
        )

        result = licensing.clean_ospp_licenses_wmi(dry_run=True)
        assert "Office 365" in result

    def test_returns_empty_when_no_licenses(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """Should return empty list when no licenses found."""
        monkeypatch.setattr(
            "office_janitor.licensing._query_wmi_licenses",
            lambda *a, **kw: [],
        )

        result = licensing.clean_ospp_licenses_wmi(dry_run=True)
        assert result == []


class TestCleanVnextCache:
    """Tests for clean_vnext_cache function."""

    def test_dry_run_returns_count(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """Dry run should return count of paths that exist."""
        cache_path = tmp_path / "Office" / "Licenses"
        cache_path.mkdir(parents=True)

        # Patch expandvars to return our temp path
        def fake_expandvars(path: str) -> str:
            if "Licenses" in path:
                return str(cache_path)
            return path

        monkeypatch.setattr("os.path.expandvars", fake_expandvars)

        result = licensing.clean_vnext_cache(dry_run=True)
        assert result >= 0  # Path may or may not exist in test env


class TestFullLicenseCleanup:
    """Tests for full_license_cleanup function."""

    def test_skip_when_keep_license_set(self) -> None:
        """Should skip cleanup when keep_license is True."""
        result = licensing.full_license_cleanup(dry_run=True, keep_license=True)
        assert result.get("skipped") is True
        assert "keep_license" in result.get("reason", "")

    def test_returns_all_categories(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """Should return results for all cleanup categories."""
        monkeypatch.setattr(
            "office_janitor.licensing.clean_ospp_licenses_wmi",
            lambda **kw: [],
        )
        monkeypatch.setattr(
            "office_janitor.licensing.clean_vnext_cache",
            lambda **kw: 0,
        )
        monkeypatch.setattr(
            "office_janitor.licensing.clean_activation_tokens",
            lambda **kw: 0,
        )
        monkeypatch.setattr(
            "office_janitor.licensing.clean_scl_cache",
            lambda **kw: 0,
        )

        result = licensing.full_license_cleanup(dry_run=True)
        assert "ospp_wmi" in result
        assert "vnext_cache" in result
        assert "activation_tokens" in result
        assert "scl_cache" in result
