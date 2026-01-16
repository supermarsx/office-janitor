"""!
@file test_appx_uninstall.py
@brief Tests for Microsoft Store (AppX) Office package removal.
"""

from __future__ import annotations

import subprocess
from unittest import mock


class TestDetectOfficeAppxPackages:
    """Tests for detect_office_appx_packages function."""

    def test_returns_empty_list_when_no_packages(self) -> None:
        """Should return empty list when no Office AppX packages found."""
        from office_janitor import appx_uninstall

        with mock.patch.object(appx_uninstall, "_run_powershell") as mock_ps:
            mock_ps.return_value = subprocess.CompletedProcess(
                args=[], returncode=0, stdout="", stderr=""
            )
            result = appx_uninstall.detect_office_appx_packages()
            assert result == []

    def test_returns_packages_when_found(self) -> None:
        """Should return package info when Office AppX packages found."""
        from office_janitor import appx_uninstall

        mock_json = (
            '{"Name":"Microsoft.Office.Desktop",'
            '"PackageFullName":"Microsoft.Office.Desktop_16.0.0.0_x64",'
            '"Version":"16.0.0.0"}'
        )

        with mock.patch.object(appx_uninstall, "_run_powershell") as mock_ps:
            mock_ps.return_value = subprocess.CompletedProcess(
                args=[], returncode=0, stdout=mock_json, stderr=""
            )
            result = appx_uninstall.detect_office_appx_packages()
            # May have multiple matches from different patterns
            assert any(pkg.get("Name") == "Microsoft.Office.Desktop" for pkg in result)

    def test_handles_multiple_packages(self) -> None:
        """Should handle multiple packages in JSON array."""
        from office_janitor import appx_uninstall

        mock_json = (
            '[{"Name":"Microsoft.Office.Desktop.Excel",'
            '"PackageFullName":"Excel_1.0","Version":"1.0"},'
            '{"Name":"Microsoft.Office.Desktop.Word",'
            '"PackageFullName":"Word_1.0","Version":"1.0"}]'
        )

        with mock.patch.object(appx_uninstall, "_run_powershell") as mock_ps:
            mock_ps.return_value = subprocess.CompletedProcess(
                args=[], returncode=0, stdout=mock_json, stderr=""
            )
            result = appx_uninstall.detect_office_appx_packages()
            names = [pkg.get("Name") for pkg in result]
            assert "Microsoft.Office.Desktop.Excel" in names
            assert "Microsoft.Office.Desktop.Word" in names

    def test_deduplicates_packages(self) -> None:
        """Should deduplicate packages by PackageFullName."""
        from office_janitor import appx_uninstall

        # Simulate same package returned for multiple patterns
        mock_json = (
            '{"Name":"Microsoft.Office.Desktop","PackageFullName":"Same_1.0","Version":"1.0"}'
        )

        with mock.patch.object(appx_uninstall, "_run_powershell") as mock_ps:
            mock_ps.return_value = subprocess.CompletedProcess(
                args=[], returncode=0, stdout=mock_json, stderr=""
            )
            result = appx_uninstall.detect_office_appx_packages()
            full_names = [pkg.get("PackageFullName") for pkg in result]
            # Should only appear once even if multiple patterns match
            assert full_names.count("Same_1.0") == 1

    def test_handles_timeout(self) -> None:
        """Should handle PowerShell timeout gracefully."""
        from office_janitor import appx_uninstall

        with mock.patch.object(appx_uninstall, "_run_powershell") as mock_ps:
            mock_ps.side_effect = subprocess.TimeoutExpired(cmd="test", timeout=60)
            result = appx_uninstall.detect_office_appx_packages()
            assert result == []

    def test_handles_os_error(self) -> None:
        """Should handle OS errors gracefully."""
        from office_janitor import appx_uninstall

        with mock.patch.object(appx_uninstall, "_run_powershell") as mock_ps:
            mock_ps.side_effect = OSError("PowerShell not found")
            result = appx_uninstall.detect_office_appx_packages()
            assert result == []


class TestRemoveOfficeAppxPackages:
    """Tests for remove_office_appx_packages function."""

    def test_dry_run_does_not_remove(self) -> None:
        """Should not actually remove packages in dry-run mode."""
        from office_janitor import appx_uninstall

        with mock.patch.object(appx_uninstall, "_run_powershell") as mock_ps:
            # Return empty for detection
            mock_ps.return_value = subprocess.CompletedProcess(
                args=[], returncode=0, stdout="", stderr=""
            )
            result = appx_uninstall.remove_office_appx_packages(
                packages=["TestPackage"], dry_run=True
            )
            assert len(result) == 1
            assert result[0]["success"] is True
            assert result[0]["dry_run"] is True
            # Should not have called PowerShell for actual removal
            # (only for detection if packages=None)

    def test_removes_specified_packages(self) -> None:
        """Should remove specified packages."""
        from office_janitor import appx_uninstall

        with mock.patch.object(appx_uninstall, "_run_powershell") as mock_ps:
            mock_ps.return_value = subprocess.CompletedProcess(
                args=[], returncode=0, stdout="", stderr=""
            )
            result = appx_uninstall.remove_office_appx_packages(
                packages=["Microsoft.Office.Desktop"], dry_run=False
            )
            assert len(result) == 1
            assert result[0]["success"] is True

    def test_handles_removal_failure(self) -> None:
        """Should handle package removal failure."""
        from office_janitor import appx_uninstall

        with mock.patch.object(appx_uninstall, "_run_powershell") as mock_ps:
            mock_ps.return_value = subprocess.CompletedProcess(
                args=[], returncode=1, stdout="", stderr="Access denied"
            )
            result = appx_uninstall.remove_office_appx_packages(
                packages=["TestPackage"], dry_run=False
            )
            assert len(result) == 1
            assert result[0]["success"] is False
            assert "Access denied" in str(result[0]["error"])

    def test_returns_empty_when_no_packages(self) -> None:
        """Should return empty list when no packages to remove."""
        from office_janitor import appx_uninstall

        with mock.patch.object(appx_uninstall, "detect_office_appx_packages") as mock_detect:
            mock_detect.return_value = []
            result = appx_uninstall.remove_office_appx_packages(packages=None, dry_run=False)
            assert result == []


class TestRemoveProvisionedAppxPackages:
    """Tests for remove_provisioned_appx_packages function."""

    def test_dry_run_does_not_remove(self) -> None:
        """Should not actually remove provisioned packages in dry-run mode."""
        from office_janitor import appx_uninstall

        mock_json = (
            '{"DisplayName":"Microsoft.Office.Desktop",'
            '"PackageName":"Microsoft.Office.Desktop_1.0"}'
        )

        with mock.patch.object(appx_uninstall, "_run_powershell") as mock_ps:
            mock_ps.return_value = subprocess.CompletedProcess(
                args=[], returncode=0, stdout=mock_json, stderr=""
            )
            result = appx_uninstall.remove_provisioned_appx_packages(dry_run=True)
            # Should find and report the package
            assert len(result) >= 1
            assert result[0]["dry_run"] is True
            assert result[0]["success"] is True

    def test_returns_empty_when_no_provisioned_packages(self) -> None:
        """Should return empty list when no provisioned packages found."""
        from office_janitor import appx_uninstall

        with mock.patch.object(appx_uninstall, "_run_powershell") as mock_ps:
            mock_ps.return_value = subprocess.CompletedProcess(
                args=[], returncode=0, stdout="", stderr=""
            )
            result = appx_uninstall.remove_provisioned_appx_packages(dry_run=False)
            assert result == []


class TestIsOfficeStoreInstall:
    """Tests for is_office_store_install function."""

    def test_returns_true_when_packages_found(self) -> None:
        """Should return True when Office AppX packages are detected."""
        from office_janitor import appx_uninstall

        with mock.patch.object(appx_uninstall, "detect_office_appx_packages") as mock_detect:
            mock_detect.return_value = [{"Name": "Microsoft.Office.Desktop"}]
            assert appx_uninstall.is_office_store_install() is True

    def test_returns_false_when_no_packages(self) -> None:
        """Should return False when no Office AppX packages are detected."""
        from office_janitor import appx_uninstall

        with mock.patch.object(appx_uninstall, "detect_office_appx_packages") as mock_detect:
            mock_detect.return_value = []
            assert appx_uninstall.is_office_store_install() is False


class TestGetAppxPackageInfo:
    """Tests for get_appx_package_info function."""

    def test_returns_package_info(self) -> None:
        """Should return package info for existing package."""
        from office_janitor import appx_uninstall

        mock_json = '{"Name":"Microsoft.Office.Desktop","Version":"16.0.0.0","Architecture":"X64"}'

        with mock.patch.object(appx_uninstall, "_run_powershell") as mock_ps:
            mock_ps.return_value = subprocess.CompletedProcess(
                args=[], returncode=0, stdout=mock_json, stderr=""
            )
            result = appx_uninstall.get_appx_package_info("Microsoft.Office.Desktop")
            assert result is not None
            assert result["Name"] == "Microsoft.Office.Desktop"

    def test_returns_none_when_not_found(self) -> None:
        """Should return None when package not found."""
        from office_janitor import appx_uninstall

        with mock.patch.object(appx_uninstall, "_run_powershell") as mock_ps:
            mock_ps.return_value = subprocess.CompletedProcess(
                args=[], returncode=0, stdout="", stderr=""
            )
            result = appx_uninstall.get_appx_package_info("NonExistent")
            assert result is None
