"""!
@brief Validate uninstall helper command composition and retry behaviour.
@details Ensures :mod:`msi_uninstall` and :mod:`c2r_uninstall` build the
expected ``msiexec``/Click-to-Run commands, honour dry-run semantics, and
surface failures through informative exceptions.
"""

from __future__ import annotations

import pathlib
import sys
from pathlib import Path

import pytest

PROJECT_ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import (  # noqa: E402
    c2r_integrator,
    c2r_odt,
    c2r_uninstall,
    command_runner,
    logging_ext,
    msi_uninstall,
)


def _command_result(
    command: list[str], returncode: int = 0, *, skipped: bool = False
) -> command_runner.CommandResult:
    """!
    @brief Convenience factory for :class:`CommandResult` instances.
    """

    return command_runner.CommandResult(
        command=command,
        returncode=returncode,
        stdout="",
        stderr="",
        duration=0.1,
        skipped=skipped,
    )


def test_msi_uninstall_builds_msiexec_command(monkeypatch, tmp_path) -> None:
    """!
    @brief Ensure ``msiexec`` commands are constructed and verification runs.
    """

    logging_ext.setup_logging(tmp_path)
    executed: list[list[str]] = []
    state = {"present": True}

    def fake_run_command(
        command: list[str],
        *,
        event: str,
        timeout: float | None = None,
        dry_run: bool = False,
        human_message: str | None = None,
        extra: dict | None = None,
    ) -> command_runner.CommandResult:
        executed.append(command)
        state["present"] = False
        return _command_result(command)

    monkeypatch.setattr(msi_uninstall.command_runner, "run_command", fake_run_command)
    monkeypatch.setattr(
        msi_uninstall.registry_tools, "key_exists", lambda *_, **__: state["present"]
    )
    monkeypatch.setattr(msi_uninstall.time, "sleep", lambda *_: None)

    record = {
        "product": "Office",
        "product_code": "{91160000-0011-0000-0000-0000000FF1CE}",
        "uninstall_handles": [
            "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{91160000-0011-0000-0000-0000000FF1CE}"
        ],
    }
    msi_uninstall.uninstall_products([record])

    assert executed, "Expected msiexec to be invoked"
    command = executed[0]
    assert command[0].lower().endswith("msiexec.exe")
    assert command[1] == "/x"
    assert command[2].startswith("{")
    assert "/qb!" in command
    assert "/norestart" in command


def test_msi_uninstall_falls_back_to_setup_when_msiexec_fails(monkeypatch, tmp_path) -> None:
    """!
    @brief Setup-based maintenance executables should be attempted after ``msiexec`` failures.
    """

    logging_ext.setup_logging(tmp_path)
    executed: list[tuple[list[str], str, bool]] = []
    probe_states: list[bool] = []
    state = {"present": True}

    setup_path = tmp_path / "setup.exe"
    setup_path.write_text("dummy")

    def fake_run_command(
        command: list[str],
        *,
        event: str,
        timeout: float | None = None,
        dry_run: bool = False,
        human_message: str | None = None,
        extra: dict | None = None,
    ) -> command_runner.CommandResult:
        executed.append((command, event, dry_run))
        if command[0].lower().endswith("msiexec.exe"):
            return _command_result(command, returncode=1603)
        state["present"] = False
        return _command_result(command)

    def fake_key_exists(*_: object, **__: object) -> bool:
        probe_states.append(state["present"])
        return state["present"]

    monkeypatch.setattr(msi_uninstall.command_runner, "run_command", fake_run_command)
    monkeypatch.setattr(msi_uninstall.registry_tools, "key_exists", fake_key_exists)
    monkeypatch.setattr(msi_uninstall.time, "sleep", lambda *_: None)

    record = {
        "product": "Office",
        "product_code": "{91160000-0011-0000-0000-0000000FF1CE}",
        "maintenance_paths": [str(setup_path)],
        "properties": {
            "maintenance_paths": [str(setup_path)],
            "display_icon": f"{setup_path},0",
        },
    }

    msi_uninstall.uninstall_products([record])

    assert len(executed) >= 2, "Expected setup.exe fallback after msiexec failure"
    msiexec_events = [item for item in executed if item[0][0].lower().endswith("msiexec.exe")]
    setup_events = [item for item in executed if item[0][0] == str(setup_path)]
    assert msiexec_events, "msiexec should run before setup fallback"
    assert setup_events, "setup.exe fallback should execute"
    msiexec_command, msiexec_event, _ = msiexec_events[0]
    setup_command, setup_event, dry_run = setup_events[-1]
    assert msiexec_event == "msi_uninstall"
    # setup.exe command now includes product code: [setup.exe, /uninstall, {GUID}]
    assert "/uninstall" in setup_command
    assert setup_event == "msi_setup_uninstall"
    assert dry_run is False
    assert False in probe_states, "Verification should observe removal state"


def test_msi_uninstall_dry_run(monkeypatch, tmp_path) -> None:
    """!
    @brief Dry-run mode should record the plan without executing ``msiexec``.
    """

    logging_ext.setup_logging(tmp_path)
    called = False

    def fake_run_command(
        command: list[str],
        *,
        event: str,
        timeout: float | None = None,
        dry_run: bool = False,
        human_message: str | None = None,
        extra: dict | None = None,
    ) -> command_runner.CommandResult:
        nonlocal called
        called = True
        assert dry_run is True
        return _command_result(command, skipped=True)

    monkeypatch.setattr(msi_uninstall.command_runner, "run_command", fake_run_command)
    monkeypatch.setattr(msi_uninstall.registry_tools, "key_exists", lambda *_, **__: True)

    msi_uninstall.uninstall_products(
        [
            {
                "product": "Office",
                "product_code": "{91160000-0011-0000-0000-0000000FF1CE}",
            }
        ],
        dry_run=True,
    )

    assert called, "Dry-run should still build the command"


def test_msi_uninstall_reports_failure(monkeypatch, tmp_path) -> None:
    """!
    @brief Non-zero return codes propagate as ``RuntimeError`` instances.
    """

    logging_ext.setup_logging(tmp_path)

    def fake_run_command(
        command: list[str],
        *,
        event: str,
        timeout: float | None = None,
        dry_run: bool = False,
        human_message: str | None = None,
        extra: dict | None = None,
    ) -> command_runner.CommandResult:
        return _command_result(command, returncode=1603)

    monkeypatch.setattr(msi_uninstall.command_runner, "run_command", fake_run_command)
    monkeypatch.setattr(msi_uninstall.registry_tools, "key_exists", lambda *_, **__: True)

    with pytest.raises(RuntimeError) as excinfo:
        msi_uninstall.uninstall_products(["{BAD-CODE}"])

    assert "BAD-CODE" in str(excinfo.value)


def test_msi_uninstall_prompts_and_retries_when_installer_busy(monkeypatch, tmp_path) -> None:
    """!
    @brief ``ERROR_INSTALL_ALREADY_RUNNING`` should trigger guidance and retries.
    """

    logging_ext.setup_logging(tmp_path)
    executed: list[list[str]] = []
    prompts: list[str] = []
    sleeps: list[float] = []
    state = {"present": True, "attempt": 0}

    def fake_run_command(
        command: list[str],
        *,
        event: str,
        timeout: float | None = None,
        dry_run: bool = False,
        human_message: str | None = None,
        extra: dict | None = None,
    ) -> command_runner.CommandResult:
        executed.append(command)
        state["attempt"] += 1
        if state["attempt"] == 1:
            return _command_result(command, returncode=msi_uninstall.MSI_BUSY_RETURN_CODE)
        state["present"] = False
        return _command_result(command)

    def fake_input(prompt: str) -> str:
        prompts.append(prompt)
        return "y"

    monkeypatch.setattr(msi_uninstall.command_runner, "run_command", fake_run_command)
    monkeypatch.setattr(
        msi_uninstall.registry_tools, "key_exists", lambda *_, **__: state["present"]
    )
    monkeypatch.setattr(msi_uninstall.time, "sleep", lambda seconds: sleeps.append(seconds))

    record = {
        "product": "Office",
        "product_code": "{91160000-0011-0000-0000-0000000FF1CE}",
        "uninstall_handles": [
            "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\{91160000-0011-0000-0000-0000000FF1CE}"
        ],
    }

    msi_uninstall.uninstall_products([record], retries=2, busy_input_func=fake_input)

    assert len(executed) == 2, "Expected retry after busy installer"
    assert prompts, "Operator prompt should be displayed for busy installer"
    assert sleeps and sleeps[0] == pytest.approx(msi_uninstall.MSI_RETRY_DELAY)


def test_msi_uninstall_cancels_when_busy_operator_declines(monkeypatch, tmp_path) -> None:
    """!
    @brief Operator refusal should surface busy failures without additional retries.
    """

    logging_ext.setup_logging(tmp_path)
    executed: list[list[str]] = []

    def fake_run_command(
        command: list[str],
        *,
        event: str,
        timeout: float | None = None,
        dry_run: bool = False,
        human_message: str | None = None,
        extra: dict | None = None,
    ) -> command_runner.CommandResult:
        executed.append(command)
        return _command_result(command, returncode=msi_uninstall.MSI_BUSY_RETURN_CODE)

    monkeypatch.setattr(msi_uninstall.command_runner, "run_command", fake_run_command)
    monkeypatch.setattr(msi_uninstall.registry_tools, "key_exists", lambda *_, **__: True)

    def fail_sleep(*_, **__):
        raise AssertionError("sleep not expected")

    monkeypatch.setattr(msi_uninstall.time, "sleep", fail_sleep)

    record = {
        "product": "Office",
        "product_code": "{91160000-0011-0000-0000-0000000FF1CE}",
    }

    with pytest.raises(RuntimeError) as excinfo:
        msi_uninstall.uninstall_products([record], retries=2, busy_input_func=lambda *_: "n")

    assert "91160000-0011-0000-0000-0000000FF1CE" in str(excinfo.value)
    assert len(executed) == 1, "Should not retry when operator declines"


def test_c2r_uninstall_prefers_client(monkeypatch, tmp_path) -> None:
    """!
    @brief OfficeC2RClient.exe should be preferred when available.
    """

    logging_ext.setup_logging(tmp_path)
    executed: list[tuple[list[str], dict]] = []
    state = {"present": True}

    def fake_run_command(
        command: list[str],
        *,
        event: str,
        timeout: float | None = None,
        dry_run: bool = False,
        human_message: str | None = None,
        extra: dict | None = None,
    ) -> command_runner.CommandResult:
        executed.append((command, extra or {}))
        state["present"] = False
        return _command_result(command)

    monkeypatch.setattr(c2r_uninstall.command_runner, "run_command", fake_run_command)
    monkeypatch.setattr(
        c2r_uninstall.tasks_services, "stop_services", lambda services, *, timeout=30: None
    )
    monkeypatch.setattr(c2r_uninstall, "_handles_present", lambda target: state["present"])
    monkeypatch.setattr(c2r_uninstall, "_install_paths_present", lambda target: state["present"])
    monkeypatch.setattr(c2r_uninstall.time, "sleep", lambda *_: None)

    def fake_find_existing_path(candidates):
        for candidate in candidates:
            try:
                if candidate.name.lower() == "officec2rclient.exe" and candidate.exists():
                    return candidate
            except OSError:
                continue
        for candidate in candidates:
            try:
                if candidate.name.lower() == "setup.exe" and candidate.exists():
                    return candidate
            except OSError:
                continue
        return None

    monkeypatch.setattr(c2r_uninstall, "_find_existing_path", fake_find_existing_path)

    client_path = tmp_path / "OfficeC2RClient.exe"
    client_path.write_text("")

    config = {
        "product": "Microsoft 365 Apps",
        "release_ids": ["O365ProPlusRetail"],
        "client_paths": [client_path],
        "uninstall_handles": [
            "HKLM\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\ProductReleaseIDs\\O365ProPlusRetail"
        ],
        "install_path": str(tmp_path),
    }

    c2r_uninstall.uninstall_products(config)

    assert executed, "Expected OfficeC2RClient.exe invocation"
    command, metadata = executed[0]
    assert command[0].endswith("OfficeC2RClient.exe")
    for arg in c2r_uninstall.C2R_CLIENT_ARGS:
        assert arg in command
    assert metadata.get("executable", "").endswith("OfficeC2RClient.exe")


def test_c2r_uninstall_fallback_to_setup(monkeypatch, tmp_path) -> None:
    """!
    @brief ``setup.exe`` fallback should run when the client is missing.
    """

    logging_ext.setup_logging(tmp_path)
    executed: list[list[str]] = []
    state = {"present": True}

    def fake_run_command(
        command: list[str],
        *,
        event: str,
        timeout: float | None = None,
        dry_run: bool = False,
        human_message: str | None = None,
        extra: dict | None = None,
    ) -> command_runner.CommandResult:
        executed.append(command)
        state["present"] = False
        return _command_result(command)

    monkeypatch.setattr(c2r_uninstall.command_runner, "run_command", fake_run_command)
    monkeypatch.setattr(
        c2r_uninstall.tasks_services, "stop_services", lambda services, *, timeout=30: None
    )
    monkeypatch.setattr(c2r_uninstall, "_handles_present", lambda target: state["present"])
    monkeypatch.setattr(c2r_uninstall, "_install_paths_present", lambda target: state["present"])
    monkeypatch.setattr(c2r_uninstall.time, "sleep", lambda *_: None)

    def fake_find_existing_path(candidates):
        for candidate in candidates:
            try:
                if candidate.name.lower() == "setup.exe" and candidate.exists():
                    return candidate
            except OSError:
                continue
        return None

    monkeypatch.setattr(c2r_uninstall, "_find_existing_path", fake_find_existing_path)

    setup_path = tmp_path / "setup.exe"
    setup_path.write_text("")

    config = {
        "release_ids": ["O365ProPlusRetail", "VisioProRetail"],
        "setup_paths": [setup_path],
        "install_path": str(tmp_path),
    }

    c2r_uninstall.uninstall_products(config)

    assert executed, "Expected setup.exe invocation"
    assert all(cmd[0].endswith("setup.exe") for cmd in executed)
    assert {cmd[2] for cmd in executed} == {"O365ProPlusRetail", "VisioProRetail"}


def test_c2r_uninstall_dry_run(monkeypatch, tmp_path) -> None:
    """!
    @brief Dry-run should not stop services or verify removal.
    """

    logging_ext.setup_logging(tmp_path)
    executed: list[list[str]] = []

    def fake_run_command(
        command: list[str],
        *,
        event: str,
        timeout: float | None = None,
        dry_run: bool = False,
        human_message: str | None = None,
        extra: dict | None = None,
    ) -> command_runner.CommandResult:
        executed.append(command)
        return _command_result(command, skipped=True)

    monkeypatch.setattr(c2r_uninstall.command_runner, "run_command", fake_run_command)

    def fail_if_called(*args, **kwargs):
        raise AssertionError("Services should not stop in dry-run")

    monkeypatch.setattr(c2r_uninstall.tasks_services, "stop_services", fail_if_called)
    monkeypatch.setattr(c2r_uninstall.registry_tools, "key_exists", lambda *_, **__: True)

    client_path = tmp_path / "OfficeC2RClient.exe"
    client_path.write_text("")

    c2r_uninstall.uninstall_products(
        {
            "release_ids": ["O365ProPlusRetail"],
            "client_paths": [client_path],
        },
        dry_run=True,
    )

    assert executed, "Dry-run should still build a command"


# ---------------------------------------------------------------------------
# Tests for ODT (Office Deployment Tool) functions
# ---------------------------------------------------------------------------


class TestBuildRemoveXml:
    """Tests for build_remove_xml function."""

    def test_creates_xml_file(self, tmp_path) -> None:
        """Should create an XML file with removal config."""
        output = tmp_path / "RemoveAll.xml"
        result = c2r_uninstall.build_remove_xml(output, quiet=True)

        assert result == output
        assert output.exists()
        content = output.read_text()
        assert "<Configuration>" in content
        assert "Remove All" in content
        assert "TRUE" in content

    def test_quiet_mode_uses_none_level(self, tmp_path) -> None:
        """Quiet mode should use Display Level None."""
        output = tmp_path / "quiet.xml"
        c2r_uninstall.build_remove_xml(output, quiet=True)

        content = output.read_text()
        assert 'Level="None"' in content

    def test_interactive_mode_uses_full_level(self, tmp_path) -> None:
        """Non-quiet mode should use Display Level Full."""
        output = tmp_path / "interactive.xml"
        c2r_uninstall.build_remove_xml(output, quiet=False)

        content = output.read_text()
        assert 'Level="Full"' in content


class TestOdtDownloadUrls:
    """Tests for ODT download URL constants."""

    def test_urls_defined_for_common_versions(self) -> None:
        """Should have URLs for Office 15 and 16."""
        assert 16 in c2r_uninstall.ODT_DOWNLOAD_URLS
        assert 15 in c2r_uninstall.ODT_DOWNLOAD_URLS

    def test_urls_are_https_or_http(self) -> None:
        """URLs should be proper HTTP(S) URLs."""
        for _version, url in c2r_uninstall.ODT_DOWNLOAD_URLS.items():
            assert url.startswith("http://") or url.startswith("https://")


class TestDownloadOdt:
    """Tests for download_odt function."""

    def test_dry_run_returns_path_without_download(self, tmp_path) -> None:
        """Dry run should return expected path without downloading."""
        result = c2r_uninstall.download_odt(16, tmp_path, dry_run=True)

        assert result is not None
        assert result == tmp_path / "setup.exe"

    def test_returns_none_for_invalid_version(self, tmp_path) -> None:
        """Should return None for unsupported version."""
        result = c2r_uninstall.download_odt(999, tmp_path, dry_run=True)

        # Version 999 doesn't exist in ODT_DOWNLOAD_URLS
        assert result is None


class TestFindOrDownloadOdt:
    """Tests for find_or_download_odt function."""

    def test_finds_existing_setup(self, tmp_path, monkeypatch) -> None:
        """Should find local setup.exe without downloading."""
        setup_path = tmp_path / "setup.exe"
        setup_path.write_text("fake")

        # Patch the actual source module (c2r_odt) not the re-export
        monkeypatch.setattr(c2r_odt, "C2R_SETUP_CANDIDATES", (setup_path,))

        result = c2r_uninstall.find_or_download_odt()
        assert result == setup_path


class TestUninstallViaOdt:
    """Tests for uninstall_via_odt function."""

    def test_dry_run_returns_success(self, tmp_path, monkeypatch) -> None:
        """Dry run should return 0 without execution."""
        setup_path = tmp_path / "setup.exe"
        setup_path.write_text("fake")

        result = c2r_uninstall.uninstall_via_odt(setup_path, dry_run=True)
        assert result == 0

    def test_returns_error_for_missing_odt(self, tmp_path) -> None:
        """Should return error code when ODT not found."""
        result = c2r_uninstall.uninstall_via_odt(tmp_path / "missing.exe", dry_run=False)
        assert result != 0


# ---------------------------------------------------------------------------
# Tests for C2R Integrator.exe functions
# ---------------------------------------------------------------------------


class TestFindIntegratorExe:
    """Tests for find_integrator_exe function."""

    def test_returns_none_when_not_found(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """Should return None when integrator.exe not found."""
        from pathlib import Path

        # Make all candidates not exist
        monkeypatch.setattr(Path, "exists", lambda self: False)
        result = c2r_uninstall.find_integrator_exe()
        assert result is None


class TestDeleteC2rManifests:
    """Tests for delete_c2r_manifests function."""

    def test_dry_run_returns_manifests(self, tmp_path: Path) -> None:
        """Dry run should return manifest paths without deletion."""
        # Create test structure
        integration_dir = tmp_path / "root" / "Integration"
        integration_dir.mkdir(parents=True)
        manifest1 = integration_dir / "C2RManifest.xml"
        manifest2 = integration_dir / "C2RManifest.en-us.xml"
        manifest1.write_text("<xml/>")
        manifest2.write_text("<xml/>")

        result = c2r_uninstall.delete_c2r_manifests(tmp_path, dry_run=True)

        assert len(result) == 2
        assert manifest1.exists()  # Still exists (dry run)
        assert manifest2.exists()

    def test_deletes_manifests(self, tmp_path: Path) -> None:
        """Should delete manifest files when not dry run."""
        integration_dir = tmp_path / "root" / "Integration"
        integration_dir.mkdir(parents=True)
        manifest = integration_dir / "C2RManifest.xml"
        manifest.write_text("<xml/>")

        result = c2r_uninstall.delete_c2r_manifests(tmp_path, dry_run=False)

        assert len(result) == 1
        assert not manifest.exists()

    def test_returns_empty_for_missing_folder(self, tmp_path: Path) -> None:
        """Should return empty list if Integration folder doesn't exist."""
        result = c2r_uninstall.delete_c2r_manifests(tmp_path, dry_run=False)
        assert result == []


class TestUnregisterC2rIntegration:
    """Tests for unregister_c2r_integration function."""

    def test_dry_run_logs_command(self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
        """Dry run should log what would be executed."""
        # Create a fake integrator
        integrator = tmp_path / "integrator.exe"
        integrator.write_bytes(b"fake")

        # Patch the actual source module (c2r_integrator) not the re-export
        monkeypatch.setattr(
            c2r_integrator,
            "find_integrator_exe",
            lambda: integrator,
        )

        result = c2r_uninstall.unregister_c2r_integration(
            tmp_path,
            "{00000000-0000-0000-0000-000000000000}",
            dry_run=True,
        )
        assert result == 0

    def test_returns_negative_when_integrator_not_found(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        """Should return -1 when integrator.exe not found."""
        from pathlib import Path

        monkeypatch.setattr(Path, "exists", lambda self: False)

        result = c2r_uninstall.unregister_c2r_integration(
            tmp_path,
            "{00000000-0000-0000-0000-000000000000}",
            dry_run=False,
        )
        assert result == -1


class TestFindC2rPackageGuids:
    """Tests for find_c2r_package_guids function."""

    def test_returns_empty_when_no_config(self, monkeypatch: pytest.MonkeyPatch) -> None:
        """Should return empty list when no C2R config found."""
        import winreg

        def fake_open_key(*args, **kwargs):
            raise FileNotFoundError("Key not found")

        monkeypatch.setattr(winreg, "OpenKey", fake_open_key)

        result = c2r_uninstall.find_c2r_package_guids()
        assert result == []
