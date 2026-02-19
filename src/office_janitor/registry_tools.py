"""!
@brief Registry management helpers.
@details Implements the registry wrappers, filters, and guard-railed backup
and delete routines described in :mod:`spec.md`. The helpers abstract winreg
view handling so callers can reason about 32-bit and 64-bit hives uniformly.
All operations honour the dry-run and safety constraints enforced by
``office_janitor.safety``.

Office-specific detection and filtering functions are in :mod:`registry_office`.
Windows Installer cleanup functions are in :mod:`registry_wi_cleanup`.
User hive management and vNext cleanup are in :mod:`registry_user`.
"""

from __future__ import annotations

import logging
import re
import shutil
from collections.abc import Iterable, Iterator, Mapping
from contextlib import contextmanager
from pathlib import Path
from typing import TYPE_CHECKING, Any

from . import exec_utils, safety, spinner

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
_WINREG_KEY_READ: int = getattr(winreg, "KEY_READ", 0)
_WINREG_HKLM: int = getattr(winreg, "HKEY_LOCAL_MACHINE", 0)
_WINREG_HKCU: int = getattr(winreg, "HKEY_CURRENT_USER", 0)
_WINREG_HKU: int = getattr(winreg, "HKEY_USERS", 0)
_WINREG_HKCR: int = getattr(winreg, "HKEY_CLASSES_ROOT", 0)

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
    @details Preserves the original case of the path portion while
    canonicalizing the hive prefix (HKLM, HKCU, etc.).
    """
    cleaned = key.strip().replace("/", "\\")
    while "\\\\" in cleaned:
        cleaned = cleaned.replace("\\\\", "\\")
    cleaned = cleaned.rstrip("\\")
    upper = cleaned.upper()
    for prefix, canonical in _CANONICAL_PREFIXES.items():
        if upper.startswith(prefix):
            # Preserve original case for the path portion after the hive prefix
            suffix = cleaned[len(prefix) :].lstrip("\\")
            return f"{canonical}\\{suffix}" if suffix else canonical
    return cleaned  # Return original case if no prefix match


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
    @details Whitelist is checked first so more specific allowed paths take
    precedence over broader blacklist entries.
    """
    normalized = _normalize_for_comparison(key)
    # Check whitelist first (more specific rules take precedence)
    if any(normalized.startswith(allowed) for allowed in _ALLOWED_REGISTRY_PREFIXES):
        return True
    # Then check blacklist
    if any(normalized.startswith(blocked) for blocked in _BLOCKED_REGISTRY_PREFIXES):
        return False
    return False


def _validate_registry_keys(keys: Iterable[str]) -> list[str]:
    """!
    @brief Ensure all keys fall within the allowed registry scope.
    @details Skips invalid keys instead of failing completely to improve resilience.
    """
    _logger = logging.getLogger(__name__)
    canonical_keys: list[str] = []
    for key in keys:
        try:
            canonical = _normalize_registry_key(key)
            if not _is_registry_path_allowed(canonical):
                _logger.warning("Skipping non-whitelisted registry key: %s", key)
                continue
            canonical_keys.append(canonical)
        except Exception as exc:
            _logger.warning("Skipping invalid registry key %s: %s", key, exc)
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
    mask = access if access is not None else _WINREG_KEY_READ
    last_error: Exception | None = None

    for candidate in _iter_access_masks(mask, view):
        try:
            handle = winreg.OpenKey(root, path, 0, candidate)
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
                winreg.CloseKey(handle)
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

    for candidate in _iter_access_masks(_WINREG_KEY_READ, view):
        try:
            handle = winreg.OpenKey(root, path, 0, candidate)
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
                        name = winreg.EnumKey(handle, index)
                    except OSError:
                        break
                    index += 1
                    if name in yielded:
                        continue
                    yielded.add(name)
                    yield name
            finally:
                winreg.CloseKey(handle)

    if not found:
        raise FileNotFoundError(path)


def iter_values(root: int, path: str, *, view: str | None = None) -> Iterator[tuple[str, Any]]:
    """!
    @brief Yield value name/value pairs for ``root``/``path`` across WOW64 views.
    """
    _ensure_winreg()
    seen: set[str] = set()
    found = False

    for candidate in _iter_access_masks(_WINREG_KEY_READ, view):
        try:
            handle = winreg.OpenKey(root, path, 0, candidate)
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
                        name, value, _ = winreg.EnumValue(handle, index)
                    except OSError:
                        break
                    index += 1
                    if name in seen:
                        continue
                    seen.add(name)
                    yield name, value
            finally:
                winreg.CloseKey(handle)

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
            value, _ = winreg.QueryValueEx(handle, value_name)
            return value
    except FileNotFoundError:
        return default
    except OSError:
        return default


def _parse_registry_path(full_path: str) -> tuple[int, str]:
    """!
    @brief Parse a full registry path string into hive constant and subpath.
    @param full_path Full path like "HKLM\\SOFTWARE\\Microsoft\\...".
    @returns Tuple of (hive_constant, subpath).
    @raises ValueError If the hive prefix is unrecognized.
    """
    normalized = _normalize_registry_key(full_path)
    if "\\" not in normalized:
        raise ValueError(f"Invalid registry path (no separator): {full_path}")

    prefix, _, subpath = normalized.partition("\\")
    hive_map = {
        "HKLM": _WINREG_HKLM,
        "HKCU": _WINREG_HKCU,
        "HKU": _WINREG_HKU,
        "HKCR": _WINREG_HKCR,
    }
    hive = hive_map.get(prefix)
    if hive is None:
        raise ValueError(f"Unrecognized registry hive: {prefix}")
    return hive, subpath


def key_exists(
    root_or_path: int | str,
    path: str | None = None,
    *,
    view: str | None = None,
) -> bool:
    """!
    @brief Determine whether the given key exists for the requested view.
    @details Accepts either (root: int, path: str) or a single full path string
        like "HKLM\\SOFTWARE\\...".
    """
    if isinstance(root_or_path, str):
        # Full path string form
        try:
            root, subpath = _parse_registry_path(root_or_path)
        except ValueError:
            return False
    else:
        # (root, path) form
        root = root_or_path
        subpath = path or ""

    try:
        _ensure_winreg()
        with open_key(root, subpath, view=view):
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
        _WINREG_HKLM: "HKLM",
        _WINREG_HKCU: "HKCU",
        _WINREG_HKU: "HKU",
        _WINREG_HKCR: "HKCR",
    }
    return mapping.get(root, hex(root))


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
    if not canonical_keys:
        logger.debug("No valid registry keys to export after validation")
        return []
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
            try:
                exec_utils.run_command(
                    [reg_executable, "export", key, str(export_path), "/y"],
                    event="registry_export",
                    dry_run=dry_run,
                    check=True,
                    extra={"key": key, "path": str(export_path)},
                )
                exported.append(export_path)
            except Exception:
                # Export failure is non-fatal - key may not exist or access denied
                # Use spinner-aware output to avoid mangled console lines
                spinner.pause_for_output()
                logger.warning(
                    "Registry export skipped for %s (key may not exist)",
                    key,
                )
                spinner.resume_after_output()
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
    @details Continues processing remaining keys even if individual deletions fail.
    Invalid keys are automatically skipped during validation.
    """
    logger = logger or _LOGGER
    canonical_keys = _validate_registry_keys(keys)
    if not canonical_keys:
        logger.debug("No valid registry keys to delete after validation")
        return
    if not safety.should_execute_destructive_action(
        "registry key deletion",
        dry_run=dry_run,
    ):
        dry_run = True
    reg_executable = shutil.which("reg")

    for key in canonical_keys:
        logger.info(
            "Preparing registry deletion",
            extra={"action": "registry-delete", "key": key, "dry_run": dry_run},
        )
        if not reg_executable:
            continue
        try:
            exec_utils.run_command(
                [reg_executable, "delete", key, "/f"],
                event="registry_delete",
                dry_run=dry_run,
                check=True,
                extra={"key": key},
            )
        except Exception:
            # Deletion failure is non-fatal - key may not exist or access denied
            spinner.pause_for_output()
            logger.warning(
                "Registry deletion skipped for %s (key may not exist or access denied)",
                key,
            )
            spinner.resume_after_output()


# ---------------------------------------------------------------------------
# Re-exports from registry_office for backwards compatibility
# ---------------------------------------------------------------------------
# These functions have been moved to registry_office.py but are re-exported
# here to maintain the existing public API.

from .registry_office import (  # noqa: E402
    cleanup_published_components,
    decode_squished_guid,
    filter_multi_string_value,
    is_office_guid,
    iter_office_uninstall_entries,
    looks_like_office_entry,
)

# Also expose the internal decode function with the original private name
_decode_squished_guid = decode_squished_guid


# ---------------------------------------------------------------------------
# Re-exports from registry_wi_cleanup for backwards compatibility
# ---------------------------------------------------------------------------
# These functions have been moved to registry_wi_cleanup.py but are re-exported
# here to maintain the existing public API.

# ---------------------------------------------------------------------------
# Re-exports from registry_user for backwards compatibility
# ---------------------------------------------------------------------------
# These functions have been moved to registry_user.py but are re-exported
# here to maintain the existing public API.
from .registry_user import (  # noqa: E402
    cleanup_taskband_registry,
    cleanup_vnext_identity_registry,
    delete_registry_value,
    get_loaded_user_hives,
    get_user_profile_hive_paths,
    get_user_profiles_directory,
    load_user_registry_hives,
    unload_user_registry_hives,
)
from .registry_wi_cleanup import (  # noqa: E402
    WI_METADATA_PATHS,
    cleanup_orphaned_typelibs,
    cleanup_protocol_handlers,
    cleanup_shell_extensions,
    cleanup_wi_orphaned_components,
    cleanup_wi_orphaned_products,
    scan_orphaned_typelibs,
    scan_wi_metadata,
    validate_wi_metadata_key,
)

__all__ = [
    # Core registry operations
    "RegistryError",
    "delete_keys",
    "export_keys",
    "get_value",
    "hive_name",
    "iter_subkeys",
    "iter_values",
    "key_exists",
    "open_key",
    "read_values",
    # Re-exports from registry_office
    "cleanup_published_components",
    "decode_squished_guid",
    "filter_multi_string_value",
    "is_office_guid",
    "iter_office_uninstall_entries",
    "looks_like_office_entry",
    # Re-exports from registry_wi_cleanup
    "WI_METADATA_PATHS",
    "cleanup_orphaned_typelibs",
    "cleanup_protocol_handlers",
    "cleanup_shell_extensions",
    "cleanup_wi_orphaned_components",
    "cleanup_wi_orphaned_products",
    "scan_orphaned_typelibs",
    "scan_wi_metadata",
    "validate_wi_metadata_key",
    # Re-exports from registry_user
    "cleanup_taskband_registry",
    "cleanup_vnext_identity_registry",
    "delete_registry_value",
    "get_loaded_user_hives",
    "get_user_profile_hive_paths",
    "get_user_profiles_directory",
    "load_user_registry_hives",
    "unload_user_registry_hives",
]
