from __future__ import annotations

import ctypes
from types import SimpleNamespace

from office_janitor import elevation, exec_utils


def test_is_admin_true(monkeypatch):
    fake_shell32 = SimpleNamespace(IsUserAnAdmin=lambda: 1)
    monkeypatch.setattr(ctypes, "windll", SimpleNamespace(shell32=fake_shell32))
    assert elevation.is_admin() is True


def test_relaunch_as_admin(monkeypatch):
    called = {}

    def fake_shell_execute(_, verb, exe, params, directory, show):
        called["verb"] = verb
        called["exe"] = exe
        called["params"] = params
        return 42

    monkeypatch.setattr(
        ctypes, "windll", SimpleNamespace(shell32=SimpleNamespace(ShellExecuteW=fake_shell_execute))
    )
    assert elevation.relaunch_as_admin(["--dry-run"]) is True
    assert called["verb"] == "runas"
    assert "--dry-run" in called["params"]


def test_run_as_limited_user_uses_runas(monkeypatch):
    monkeypatch.setattr(elevation.shutil, "which", lambda _: r"C:\Windows\System32\runas.exe")
    captured = {}

    def fake_run_command(command, **kwargs):
        captured["command"] = command
        captured["kwargs"] = kwargs
        return exec_utils.CommandResult(
            command=command, returncode=0, stdout="", stderr="", duration=0.0
        )

    monkeypatch.setattr(exec_utils, "run_command", fake_run_command)
    elevation.run_as_limited_user(["cmd.exe", "/c", "echo", "hi"], dry_run=True)
    assert captured["command"][0].lower().endswith("runas.exe")
    assert "/trustlevel:0x20000" in captured["command"]


def test_run_as_limited_user_fallback(monkeypatch):
    monkeypatch.setattr(elevation.shutil, "which", lambda _: None)
    captured = {}

    def fake_run_command(command, **kwargs):
        captured["command"] = command
        captured["kwargs"] = kwargs
        return exec_utils.CommandResult(
            command=command, returncode=0, stdout="", stderr="", duration=0.0
        )

    monkeypatch.setattr(exec_utils, "run_command", fake_run_command)
    elevation.run_as_limited_user(["cmd.exe", "/c", "echo", "hi"], dry_run=True)
    assert captured["command"][0] == "cmd.exe"
