"""!
@brief Registry management helpers.
@details Implements the registry wrappers, filters, and guard-railed backup
and delete routines described in :mod:`spec.md`. The helpers abstract winreg
view handling so callers can reason about 32-bit and 64-bit hives uniformly
while also exposing convenience predicates for Office-specific uninstall
entries. All operations honour the dry-run and safety constraints enforced by
``office_janitor.safety``.
"""

from __future__ import annotations

import logging
import re
import shutil
from collections.abc import Iterable, Iterator, Mapping
from contextlib import contextmanager
from pathlib import Path
from typing import TYPE_CHECKING, Any, cast

from . import exec_utils, safety

if TYPE_CHECKING:  # pragma: no cover - typing only
    import winreg as _winreg
else:  # pragma: no cover - runtime fallback for non-Windows tests
    try:
        import winreg as _winreg  # type: ignore[import-not-found]
    except ImportError:
        class _WinRegStub:
            KEY_READ = 0
            KEY_WRITE = 0
            HKEY_LOCAL_MACHINE = 0x80000002
            HKEY_CURRENT_USER = 0x80000001
            HKEY_CLASSES_ROOT = 0x80000000
            HKEY_USERS = 0x80000003

            def OpenKey(self, *args, **kwargs):  # type: ignore[no-untyped-def]
                raise FileNotFoundError

        _winreg = _WinRegStub()  # type: ignore[assignment]

winreg = _winreg


_LOGGER = logging.getLogger(__name__)

_CANONICAL_PREFIXES: Mapping[str, str] = {
    "HKEY_LOCAL_MACHINE": "HKLM",
    "HKLM": "HKLM",
    "HKEY_CURRENT_USER": "HKCU",
    "HKCU": "HKCU",
    "HKEY_CLASSES_ROOT": "HKCR",
    "HKCR": "HKCR",
    "HKEY_USERS": "HKU",
    "HKU": "HKU",
}

_SAFE_FILENAME_PATTERN = re.compile(r"[^A-Z0-9._-]+")

_OFFICE_KEYWORDS = (
    "office",
    "microsoft 365",
    "365",
    "visio",
    "project",
    "onenote",
    "word",
    "excel",
    "outlook",
    "powerpoint",
)


class RegistryError(RuntimeError):
    """!
    @brief Raised when registry operations cannot be completed safely.
    """


def _ensure_winreg() -> None:
    """!
    @brief Raise an informative error when ``winreg`` is unavailable.
    """

    if winreg is None:  # pragma: no cover - simplifies non-Windows test runs.
        raise FileNotFoundError("Windows registry APIs are unavailable on this platform")


def _normalize_registry_key(key: str) -> str:
    """!
    @brief Canonicalise registry key strings to ``HKXX`` prefixes.
    """

    cleaned = key.strip().replace("/", "\\")
    while "\\\\" in cleaned:
        cleaned = cleaned.replace("\\\\", "\\")
    cleaned = cleaned.rstrip("\\")
    upper = cleaned.upper()
    for prefix, canonical in _CANONICAL_PREFIXES.items():
        if upper.startswith(prefix):
            suffix = upper[len(prefix) :].lstrip("\\")
            return f"{canonical}\\{suffix}" if suffix else canonical
    return upper


def _normalize_for_comparison(key: str) -> str:
    """!
    @brief Provide a uppercase normalisation used for safety checks.
    """

    canonical = _normalize_registry_key(key)
    return canonical.upper()


_ALLOWED_REGISTRY_PREFIXES = tuple(
    _normalize_registry_key(entry) for entry in safety.REGISTRY_WHITELIST
)
_BLOCKED_REGISTRY_PREFIXES = tuple(
    _normalize_registry_key(entry) for entry in safety.REGISTRY_BLACKLIST
)


def _is_registry_path_allowed(key: str) -> bool:
    """!
    @brief Validate the registry path against the whitelist/blacklist rules.
    """

    normalized = _normalize_for_comparison(key)
    if any(normalized.startswith(blocked) for blocked in _BLOCKED_REGISTRY_PREFIXES):
        return False
    return any(normalized.startswith(allowed) for allowed in _ALLOWED_REGISTRY_PREFIXES)


def _validate_registry_keys(keys: Iterable[str]) -> list[str]:
    """!
    @brief Ensure all keys fall within the allowed registry scope.
    """

    canonical_keys: list[str] = []
    for key in keys:
        canonical = _normalize_registry_key(key)
        if not _is_registry_path_allowed(canonical):
            raise RegistryError(f"Refusing to operate on non-whitelisted registry key: {key}")
        canonical_keys.append(canonical)
    return canonical_keys


def _iter_access_masks(access: int, view: str | None) -> Iterator[int]:
    """!
    @brief Yield candidate access masks for the requested WOW64 view.
    """

    if winreg is None:  # pragma: no cover - defensive guard for type checkers.
        return iter(())

    view = (view or "auto").lower()
    if view not in {"auto", "native", "32bit", "64bit", "both"}:
        raise ValueError(f"Unsupported registry view: {view}")

    masks: list[int] = []
    native_mask = access
    wow64_32 = getattr(winreg, "KEY_WOW64_32KEY", 0)
    wow64_64 = getattr(winreg, "KEY_WOW64_64KEY", 0)

    if view in {"native", "auto"}:
        masks.append(native_mask)
    if view in {"32bit", "both", "auto"} and wow64_32:
        masks.append(access | wow64_32)
    if view in {"64bit", "both", "auto"} and wow64_64:
        masks.append(access | wow64_64)

    seen: set[int] = set()
    for mask in masks:
        if mask in seen:
            continue
        seen.add(mask)
        yield mask


@contextmanager
def open_key(
    root: int, path: str, *, access: int | None = None, view: str | None = None
) -> Iterator[Any]:
    """!
    @brief Context manager that mirrors ``winreg.OpenKey`` with view handling.
    """

    _ensure_winreg()
    mask = access if access is not None else winreg.KEY_READ  # type: ignore[union-attr]
    last_error: Exception | None = None

    for candidate in _iter_access_masks(mask, view):
        try:
            handle = winreg.OpenKey(root, path, 0, candidate)  # type: ignore[union-attr]
        except FileNotFoundError as exc:
            last_error = exc
            continue
        except OSError as exc:  # pragma: no cover - depends on host registry state.
            last_error = exc
            continue
        else:
            try:
                yield handle
            finally:
                winreg.CloseKey(handle)  # type: ignore[union-attr]
            return

    if last_error is not None:
        raise last_error
    raise FileNotFoundError(path)


def iter_subkeys(root: int, path: str, *, view: str | None = None) -> Iterator[str]:
    """!
    @brief Yield subkey names for ``root``/``path`` across WOW64 views.
    """

    _ensure_winreg()
    yielded: set[str] = set()
    found = False

    for candidate in _iter_access_masks(winreg.KEY_READ, view):  # type: ignore[union-attr]
        try:
            handle = winreg.OpenKey(root, path, 0, candidate)  # type: ignore[union-attr]
        except FileNotFoundError:
            continue
        except OSError:  # pragma: no cover - depends on registry permissions.
            continue
        else:
            found = True
            try:
                index = 0
                while True:
                    try:
                        name = winreg.EnumKey(handle, index)  # type: ignore[union-attr]
                    except OSError:
                        break
                    index += 1
                    if name in yielded:
                        continue
                    yielded.add(name)
                    yield name
            finally:
                winreg.CloseKey(handle)  # type: ignore[union-attr]

    if not found:
        raise FileNotFoundError(path)


def iter_values(root: int, path: str, *, view: str | None = None) -> Iterator[tuple[str, Any]]:
    """!
    @brief Yield value name/value pairs for ``root``/``path`` across WOW64 views.
    """

    _ensure_winreg()
    seen: set[str] = set()
    found = False

    for candidate in _iter_access_masks(winreg.KEY_READ, view):  # type: ignore[union-attr]
        try:
            handle = winreg.OpenKey(root, path, 0, candidate)  # type: ignore[union-attr]
        except FileNotFoundError:
            continue
        except OSError:  # pragma: no cover - depends on registry permissions.
            continue
        else:
            found = True
            try:
                index = 0
                while True:
                    try:
                        name, value, _ = winreg.EnumValue(handle, index)  # type: ignore[union-attr]
                    except OSError:
                        break
                    index += 1
                    if name in seen:
                        continue
                    seen.add(name)
                    yield name, value
            finally:
                winreg.CloseKey(handle)  # type: ignore[union-attr]

    if not found:
        raise FileNotFoundError(path)


def read_values(root: int, path: str, *, view: str | None = None) -> dict[str, Any]:
    """!
    @brief Read all values beneath ``root``/``path`` into a dictionary.
    """

    values: dict[str, Any] = {}
    try:
        for name, value in iter_values(root, path, view=view):
            values.setdefault(name, value)
    except FileNotFoundError:
        return {}
    except OSError:
        return {}
    return values


def get_value(
    root: int,
    path: str,
    value_name: str,
    default: Any | None = None,
    *,
    view: str | None = None,
) -> Any | None:
    """!
    @brief Read ``value_name`` beneath ``root``/``path``.
    """

    try:
        _ensure_winreg()
        with open_key(root, path, view=view) as handle:
            value, _ = winreg.QueryValueEx(handle, value_name)  # type: ignore[union-attr]
            return value
    except FileNotFoundError:
        return default
    except OSError:
        return default


def key_exists(root: int, path: str, *, view: str | None = None) -> bool:
    """!
    @brief Determine whether the given key exists for the requested view.
    """

    try:
        _ensure_winreg()
        with open_key(root, path, view=view):
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


def looks_like_office_entry(values: Mapping[str, Any]) -> bool:
    """!
    @brief Determine whether an uninstall entry corresponds to Microsoft Office.
    @details Heuristics focus on the ``DisplayName`` and ``Publisher`` fields,
    mirroring OffScrub filters to target Office suites, Visio, Project, and
    related SKUs.
    """

    display = str(values.get("DisplayName") or "").lower()
    publisher = str(values.get("Publisher") or "").lower()
    product_code = str(values.get("ProductCode") or "").upper()

    if not display and not product_code:
        return False

    if publisher and "microsoft" not in publisher:
        return False

    for keyword in _OFFICE_KEYWORDS:
        if keyword in display:
            return True

    if product_code.startswith("{901") or product_code.startswith("{911"):
        return True

    return False


def iter_office_uninstall_entries(
    roots: Iterable[tuple[int, str]],
    *,
    view: str | None = None,
) -> Iterator[tuple[int, str, dict[str, Any]]]:
    """!
    @brief Enumerate uninstall entries that resemble Office installations.
    """

    seen_paths: set[tuple[int, str]] = set()
    for hive, base_path in roots:
        try:
            subkeys = list(iter_subkeys(hive, base_path, view=view))
        except FileNotFoundError:
            continue

        for subkey in subkeys:
            relative_path = f"{base_path}\\{subkey}"
            if (hive, relative_path) in seen_paths:
                continue
            values = read_values(hive, relative_path, view=view)
            if not values:
                continue
            if looks_like_office_entry(values):
                seen_paths.add((hive, relative_path))
                yield hive, relative_path, values


def _safe_export_filename(key: str) -> str:
    """!
    @brief Generate a filesystem-safe filename for registry exports.
    """

    canonical = _normalize_registry_key(key)
    token = canonical.replace("\\", "_").upper()
    token = _SAFE_FILENAME_PATTERN.sub("_", token)
    token = token.strip("_") or "REGISTRY_EXPORT"
    return f"{token}.reg"


def _unique_export_path(directory: Path, key: str) -> Path:
    """!
    @brief Determine a unique path for the registry export file.
    """

    base_name = _safe_export_filename(key)
    candidate = directory / base_name
    counter = 1
    while candidate.exists():
        candidate = directory / f"{base_name[:-4]}_{counter}.reg"
        counter += 1
    return candidate


def export_keys(
    keys: Iterable[str],
    destination: str | Path,
    *,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> list[Path]:
    """!
    @brief Export the provided registry keys to ``.reg`` files in ``destination``.
    @details On non-Windows systems the exports become placeholder files so unit
    tests and dry-run flows can still verify orchestration logic without access
    to the native ``reg.exe`` utility.
    """

    logger = logger or _LOGGER
    dest_path = Path(destination)
    dest_path.mkdir(parents=True, exist_ok=True)

    canonical_keys = _validate_registry_keys(keys)
    reg_executable = shutil.which("reg")
    exported: list[Path] = []

    for key in canonical_keys:
        export_path = _unique_export_path(dest_path, key)
        logger.info(
            "Preparing registry export",
            extra={
                "action": "registry-export",
                "key": key,
                "path": str(export_path),
                "dry_run": dry_run,
            },
        )
        if dry_run:
            exported.append(export_path)
            continue

        if reg_executable:
            exec_utils.run_command(
                [reg_executable, "export", key, str(export_path), "/y"],
                event="registry_export",
                dry_run=dry_run,
                check=True,
                extra={"key": key, "path": str(export_path)},
            )
            exported.append(export_path)
            continue

        export_path.write_text(
            f"; Placeholder export for {key}\n",
            encoding="utf-8",
        )
        exported.append(export_path)

    return exported


def delete_keys(
    keys: Iterable[str],
    *,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> None:
    """!
    @brief Remove registry keys while respecting dry-run safeguards.
    """

    logger = logger or _LOGGER
    canonical_keys = _validate_registry_keys(keys)
    reg_executable = shutil.which("reg")

    for key in canonical_keys:
        logger.info(
            "Preparing registry deletion",
            extra={"action": "registry-delete", "key": key, "dry_run": dry_run},
        )
        if not reg_executable:
            continue
        exec_utils.run_command(
            [reg_executable, "delete", key, "/f"],
            event="registry_delete",
            dry_run=dry_run,
            check=True,
            extra={"key": key},
        )


__all__ = [
    "RegistryError",
    "delete_keys",
    "export_keys",
    "get_value",
    "hive_name",
    "iter_office_uninstall_entries",
    "iter_subkeys",
    "iter_values",
    "key_exists",
    "looks_like_office_entry",
    "open_key",
    "read_values",
]
