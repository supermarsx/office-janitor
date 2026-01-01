from __future__ import annotations

import pathlib
import sys

PROJECT_ROOT = pathlib.Path(__file__).resolve().parents[1]
SRC_PATH = PROJECT_ROOT / "src"
if str(SRC_PATH) not in sys.path:
    sys.path.insert(0, str(SRC_PATH))

from office_janitor import constants, off_scrub_scripts  # noqa: E402


def test_build_offscrub_command_msi_defaults(tmp_path):
    cmd = off_scrub_scripts.build_offscrub_command("msi", version=None, base_directory=tmp_path)
    assert isinstance(cmd, list)
    assert cmd[0] == sys.executable
    assert any(str(tmp_path) in part for part in cmd)


def test_build_offscrub_command_c2r(tmp_path):
    cmd = off_scrub_scripts.build_offscrub_command("c2r", base_directory=tmp_path)
    assert isinstance(cmd, list)
    assert cmd[0] == sys.executable
    assert any(constants.C2R_OFFSCRUB_SCRIPT in part for part in cmd)


def test_ensure_all_offscrub_shims_writes(tmp_path):
    paths = off_scrub_scripts.ensure_all_offscrub_shims(base_directory=tmp_path)
    assert paths
    for p in paths:
        assert pathlib.Path(p).exists()
