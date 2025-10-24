"""!
@brief Registry management helpers.
@details The registry tools export hives for backup, delete targeted keys, and
provide winreg utilities used throughout detection and cleanup as outlined in
the specification.
"""
from __future__ import annotations

import shutil
import subprocess
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Dict, Iterable, Iterator, Tuple

try:  # pragma: no cover - exercised through mocks on non-Windows platforms.
    import winreg
except ImportError:  # pragma: no cover - handled gracefully during tests.
    winreg = None  # type: ignore[assignment]


def _ensure_winreg() -> None:
    """!  
    @brief Raise an informative error when ``winreg`` is unavailable.
    """

    if winreg is None:  # pragma: no cover - simplifies non-Windows test runs.
        raise FileNotFoundError("Windows registry APIs are unavailable on this platform")


def export_keys(keys: Iterable[str], destination: str) -> None:
    """!
    @brief Export the provided registry keys to ``.reg`` files in ``destination``.
    @details On non-Windows systems the exports become placeholder files so unit
    tests and dry-run flows can still verify orchestration logic without access
    to the native ``reg.exe`` utility.
    """

    dest_path = Path(destination)
    dest_path.mkdir(parents=True, exist_ok=True)
    reg_executable = shutil.which("reg")

    for key in keys:
        safe_name = key.replace("\\", "_").replace("/", "_")
        export_path = dest_path / f"{safe_name}.reg"
        if reg_executable:
            subprocess.run(
                [reg_executable, "export", key, str(export_path), "/y"],
                check=True,
            )
        else:  # pragma: no cover - depends on environment availability.
            export_path.write_text(
                f"; Placeholder export for {key}\n",
                encoding="utf-8",
            )


def delete_keys(keys: Iterable[str], *, dry_run: bool = False) -> None:
    """!
    @brief Remove registry keys while respecting dry-run safeguards.
    """

    reg_executable = shutil.which("reg")

    for key in keys:
        if dry_run or not reg_executable:
            continue
        subprocess.run([reg_executable, "delete", key, "/f"], check=True)


@contextmanager
def open_key(root: int, path: str, access: int | None = None) -> Iterator[Any]:
    """!
    @brief Context manager that mirrors ``winreg.OpenKey`` while ensuring
    handles are closed correctly.
    """

    _ensure_winreg()
    access_mask = access if access is not None else winreg.KEY_READ  # type: ignore[union-attr]
    handle = winreg.OpenKey(root, path, 0, access_mask)  # type: ignore[union-attr]
    try:
        yield handle
    finally:
        winreg.CloseKey(handle)  # type: ignore[union-attr]


def iter_subkeys(root: int, path: str) -> Iterator[str]:
    """!
    @brief Yield subkey names for ``root``/``path``.
    """

    _ensure_winreg()
    with open_key(root, path) as handle:
        subkey_count, _, _ = winreg.QueryInfoKey(handle)  # type: ignore[union-attr]
        for index in range(subkey_count):
            yield winreg.EnumKey(handle, index)  # type: ignore[union-attr]


def iter_values(root: int, path: str) -> Iterator[Tuple[str, Any]]:
    """!
    @brief Yield value name/value pairs for ``root``/``path``.
    """

    _ensure_winreg()
    with open_key(root, path) as handle:
        _, value_count, _ = winreg.QueryInfoKey(handle)  # type: ignore[union-attr]
        for index in range(value_count):
            name, value, _ = winreg.EnumValue(handle, index)  # type: ignore[union-attr]
            yield name, value


def read_values(root: int, path: str) -> Dict[str, Any]:
    """!
    @brief Read all values beneath ``root``/``path`` into a dictionary.
    """

    data: Dict[str, Any] = {}
    try:
        for name, value in iter_values(root, path):
            data[name] = value
    except FileNotFoundError:
        return {}
    except OSError:
        return {}
    return data


def get_value(root: int, path: str, value_name: str, default: Any | None = None) -> Any | None:
    """!
    @brief Read ``value_name`` beneath ``root``/``path``.
    """

    try:
        _ensure_winreg()
        with open_key(root, path) as handle:
            value, _ = winreg.QueryValueEx(handle, value_name)  # type: ignore[union-attr]
            return value
    except FileNotFoundError:
        return default
    except OSError:
        return default


def key_exists(root: int, path: str) -> bool:
    """!
    @brief Determine whether the given key exists.
    """

    try:
        _ensure_winreg()
        with open_key(root, path):
            return True
    except FileNotFoundError:
        return False
    except OSError:
        return False


def hive_name(root: int) -> str:
    """!
    @brief Provide a friendly identifier for a registry hive.
    """

    mapping = {
        getattr(winreg, "HKEY_LOCAL_MACHINE", 0): "HKLM",  # type: ignore[union-attr]
        getattr(winreg, "HKEY_CURRENT_USER", 0): "HKCU",  # type: ignore[union-attr]
        getattr(winreg, "HKEY_USERS", 0): "HKU",  # type: ignore[union-attr]
        getattr(winreg, "HKEY_CLASSES_ROOT", 0): "HKCR",  # type: ignore[union-attr]
    }
    return mapping.get(root, hex(root))


__all__ = [
    "delete_keys",
    "export_keys",
    "get_value",
    "hive_name",
    "iter_subkeys",
    "iter_values",
    "key_exists",
    "open_key",
    "read_values",
]
