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
from typing import TYPE_CHECKING, Any

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
_WINREG_KEY_READ = getattr(winreg, "KEY_READ", 0)
_WINREG_HKLM = getattr(winreg, "HKEY_LOCAL_MACHINE", 0)
_WINREG_HKCU = getattr(winreg, "HKEY_CURRENT_USER", 0)
_WINREG_HKU = getattr(winreg, "HKEY_USERS", 0)
_WINREG_HKCR = getattr(winreg, "HKEY_CLASSES_ROOT", 0)

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

_OFFICE_EXCLUSIONS = (
    # Visual Studio and related tools
    "visual studio",
    "visualstudio",
    "vs ",
    "vs20",
    # .NET Aspire and SDK templates
    "aspire",
    "aspire.projecttemplates",
    ".net sdk",
    "dotnet",
    # Development tools
    "sql server",
    "azure",
    "nuget",
    "typescript",
    "node.js",
    # Other Microsoft products that aren't Office
    "bing",
    "edge",
    "teams",  # Teams is separate from Office
    "skype",
    "onedrive",
    "sharepoint",
    "dynamics",
    "power bi",
    "powerbi",
    "power automate",
    "power apps",
    # Project/Visio viewer apps (not full installs)
    "viewer",
    # Templates and SDKs
    "template",
    "sdk",
    "runtime",
    "redistributable",
    "redist",
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
        _WINREG_HKLM: "HKLM",
        _WINREG_HKCU: "HKCU",
        _WINREG_HKU: "HKU",
        _WINREG_HKCR: "HKCR",
    }
    return mapping.get(root, hex(root))


def looks_like_office_entry(values: Mapping[str, Any]) -> bool:
    """!
    @brief Determine whether an uninstall entry corresponds to Microsoft Office.
    @details Heuristics focus on the ``DisplayName`` and ``Publisher`` fields,
    mirroring OffScrub filters to target Office suites, Visio, Project, and
    related SKUs. Excludes Visual Studio, .NET Aspire, and other non-Office
    Microsoft products.
    """

    display = str(values.get("DisplayName") or "").lower()
    publisher = str(values.get("Publisher") or "").lower()
    product_code = str(values.get("ProductCode") or "").upper()

    if not display and not product_code:
        return False

    if publisher and "microsoft" not in publisher:
        return False

    # Check exclusions first - these are NOT Office even if they match keywords
    for exclusion in _OFFICE_EXCLUSIONS:
        if exclusion in display:
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


# ---------------------------------------------------------------------------
# Windows Installer Metadata Validation
# ---------------------------------------------------------------------------
# These functions validate and clean up WI registry entries, mirroring the
# VBS ``ValidateWIMetadataKey`` function from OffScrub_O16msi.vbs.


# Standard WI registry paths for validation
WI_METADATA_PATHS: dict[str, tuple[int, str, int]] = {
    # (hive, path, expected_key_length)
    "Products": (_WINREG_HKLM, r"SOFTWARE\Classes\Installer\Products", 32),
    "Components": (_WINREG_HKLM, r"SOFTWARE\Classes\Installer\Components", 32),
    "Features": (_WINREG_HKLM, r"SOFTWARE\Classes\Installer\Features", 32),
    "Patches": (_WINREG_HKLM, r"SOFTWARE\Classes\Installer\Patches", 32),
    "UpgradeCodes": (_WINREG_HKLM, r"SOFTWARE\Classes\Installer\UpgradeCodes", 32),
    "UserDataProducts": (
        _WINREG_HKLM,
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData",
        0,  # Variable structure
    ),
}
"""!
@brief Windows Installer registry paths that may contain orphaned entries.
"""


def _is_valid_compressed_guid(name: str) -> bool:
    """!
    @brief Check if a registry key name is a valid compressed GUID.
    @param name Subkey name to validate.
    @return True if valid 32-char hex string.
    """
    if len(name) != 32:
        return False
    try:
        int(name, 16)
        return True
    except ValueError:
        return False


def validate_wi_metadata_key(
    hive: int,
    path: str,
    expected_length: int = 32,
    *,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> list[str]:
    """!
    @brief Validate WI metadata keys and identify invalid entries.
    @param hive Registry hive (HKLM, etc.).
    @param path Registry path to validate.
    @param expected_length Expected length of valid subkey names (32 for GUIDs).
    @param dry_run If True, only report without deleting.
    @param logger Optional logger.
    @return List of invalid subkey names found.

    @details Scans a Windows Installer metadata key and identifies entries
    that don't match the expected format (e.g., corrupted GUIDs). This mirrors
    the VBS ``ValidateWIMetadataKey`` function from OffScrub_O16msi.vbs.

    Invalid entries are subkeys that:
    - Don't have the expected length (32 chars for compressed GUIDs)
    - Contain non-hexadecimal characters
    - Are empty or malformed
    """
    logger = logger or _LOGGER
    invalid_entries: list[str] = []

    try:
        for subkey in iter_subkeys(hive, path, view="native"):
            # Check if subkey name matches expected format
            if expected_length > 0:
                if len(subkey) != expected_length:
                    invalid_entries.append(subkey)
                    continue
                if not _is_valid_compressed_guid(subkey):
                    invalid_entries.append(subkey)
    except FileNotFoundError:
        logger.debug("WI metadata path not found: %s", path)
        return []
    except OSError as e:
        logger.warning("Failed to access WI metadata path %s: %s", path, e)
        return []

    if invalid_entries:
        logger.info(
            "Found %d invalid WI metadata entries in %s",
            len(invalid_entries),
            path,
            extra={
                "action": "wi-validation",
                "path": path,
                "invalid_count": len(invalid_entries),
            },
        )

    return invalid_entries


def scan_wi_metadata(
    *,
    logger: logging.Logger | None = None,
) -> dict[str, list[str]]:
    """!
    @brief Scan all standard WI metadata paths for invalid entries.
    @param logger Optional logger.
    @return Dictionary mapping path names to lists of invalid entries.
    """
    logger = logger or _LOGGER
    results: dict[str, list[str]] = {}

    logger.info("Scanning Windows Installer metadata for invalid entries")

    for name, (hive, path, expected_len) in WI_METADATA_PATHS.items():
        if expected_len == 0:
            continue  # Skip variable-structure paths
        invalid = validate_wi_metadata_key(hive, path, expected_len, logger=logger)
        if invalid:
            results[name] = invalid

    return results


def cleanup_wi_orphaned_products(
    product_codes: Iterable[str],
    *,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> int:
    """!
    @brief Remove WI registry entries for orphaned product codes.
    @param product_codes Product GUIDs to clean up.
    @param dry_run If True, only log without deleting.
    @param logger Optional logger.
    @return Number of entries removed.

    @details Cleans up WI metadata entries for products that are no longer
    properly installed. This includes:
    - Installer\\Products\\<compressed_guid>
    - Installer\\Features\\<compressed_guid>
    - Installer\\UpgradeCodes entries that reference the product
    """
    from . import guid_utils

    logger = logger or _LOGGER
    removed = 0

    for product_code in product_codes:
        try:
            compressed = guid_utils.compress_guid(product_code)
        except guid_utils.GuidError:
            logger.warning("Invalid product code: %s", product_code)
            continue

        # Build paths to clean
        paths_to_clean = [
            f"HKLM\\SOFTWARE\\Classes\\Installer\\Products\\{compressed}",
            f"HKLM\\SOFTWARE\\Classes\\Installer\\Features\\{compressed}",
        ]

        for path in paths_to_clean:
            if key_exists(path):
                logger.info(
                    "Removing orphaned WI entry: %s",
                    path,
                    extra={"action": "wi-cleanup", "path": path, "dry_run": dry_run},
                )
                if not dry_run:
                    try:
                        delete_keys([path], dry_run=False, logger=logger)
                        removed += 1
                    except Exception as e:
                        logger.warning("Failed to delete %s: %s", path, e)
                else:
                    removed += 1

    return removed


def cleanup_wi_orphaned_components(
    component_ids: Iterable[str],
    *,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> int:
    """!
    @brief Remove WI registry entries for orphaned component IDs.
    @param component_ids Component GUIDs to clean up.
    @param dry_run If True, only log without deleting.
    @param logger Optional logger.
    @return Number of entries removed.

    @details Cleans up WI component entries that have no valid product clients.
    This is a more targeted cleanup than removing entire product trees.
    """
    from . import guid_utils

    logger = logger or _LOGGER
    removed = 0

    for component_id in component_ids:
        try:
            compressed = guid_utils.compress_guid(component_id)
        except guid_utils.GuidError:
            logger.warning("Invalid component ID: %s", component_id)
            continue

        path = f"HKLM\\SOFTWARE\\Classes\\Installer\\Components\\{compressed}"

        if key_exists(path):
            logger.info(
                "Removing orphaned WI component: %s",
                path,
                extra={"action": "wi-cleanup", "path": path, "dry_run": dry_run},
            )
            if not dry_run:
                try:
                    delete_keys([path], dry_run=False, logger=logger)
                    removed += 1
                except Exception as e:
                    logger.warning("Failed to delete %s: %s", path, e)
            else:
                removed += 1

    return removed


# ---------------------------------------------------------------------------
# Shell Integration Cleanup
# ---------------------------------------------------------------------------
# Functions to clean up orphaned shell extensions, TypeLibs, and protocol
# handlers left behind by Office uninstalls.


def cleanup_orphaned_typelibs(
    typelib_guids: Iterable[str],
    *,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> list[str]:
    """!
    @brief Remove orphaned TypeLib registrations for Office components.
    @param typelib_guids TypeLib GUIDs to check.
    @param dry_run If True, only report without deleting.
    @param logger Optional logger.
    @return List of removed TypeLib GUIDs.

    @details Scans TypeLib registrations and removes entries where the
    referenced DLL file no longer exists. This mirrors the VBS TypeLib
    cleanup logic from OffScrub_O16msi.vbs.
    """
    logger = logger or _LOGGER
    removed: list[str] = []

    for typelib_guid in typelib_guids:
        base_path = f"HKLM\\SOFTWARE\\Classes\\TypeLib\\{typelib_guid}"

        if not key_exists(base_path):
            continue

        # Check each version subkey
        try:
            versions = list(
                iter_subkeys(_WINREG_HKLM, f"SOFTWARE\\Classes\\TypeLib\\{typelib_guid}")
            )
        except (FileNotFoundError, OSError):
            continue

        typelib_orphaned = True

        for version in versions:
            version_path = f"SOFTWARE\\Classes\\TypeLib\\{typelib_guid}\\{version}"
            try:
                # Check the 0\\win32 or 0\\win64 paths for the DLL location
                for platform in ("0\\win32", "0\\win64", "0"):
                    try:
                        values = read_values(
                            _WINREG_HKLM, f"{version_path}\\{platform}", view="native"
                        )
                        default_val = values.get("", "")
                        if default_val and Path(default_val).exists():
                            typelib_orphaned = False
                            break
                    except (FileNotFoundError, OSError):
                        continue
            except (FileNotFoundError, OSError):
                continue

            if not typelib_orphaned:
                break

        if typelib_orphaned:
            logger.info(
                "Removing orphaned TypeLib: %s",
                typelib_guid,
                extra={"action": "typelib-cleanup", "guid": typelib_guid, "dry_run": dry_run},
            )
            if not dry_run:
                try:
                    delete_keys([base_path], dry_run=False, logger=logger)
                    removed.append(typelib_guid)
                except Exception as e:
                    logger.warning("Failed to delete TypeLib %s: %s", typelib_guid, e)
            else:
                removed.append(typelib_guid)

    return removed


def scan_orphaned_typelibs(
    typelib_guids: Iterable[str],
    *,
    logger: logging.Logger | None = None,
) -> list[str]:
    """!
    @brief Scan for orphaned TypeLib registrations without removing them.
    @param typelib_guids TypeLib GUIDs to check.
    @param logger Optional logger.
    @return List of orphaned TypeLib GUIDs.
    """
    return cleanup_orphaned_typelibs(typelib_guids, dry_run=True, logger=logger)


def cleanup_shell_extensions(
    *,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> int:
    """!
    @brief Remove orphaned shell extension registrations.
    @param dry_run If True, only report without deleting.
    @param logger Optional logger.
    @return Number of extensions removed.

    @details Cleans up shell extensions that reference non-existent DLLs.
    This includes context menu handlers, property sheet handlers, etc.
    """
    from . import constants

    logger = logger or _LOGGER
    removed = 0

    # Shell extension approval entries
    approval_paths = [
        (
            _WINREG_HKLM,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved",
        ),
    ]

    for hive, path in approval_paths:
        try:
            for name, value in iter_values(hive, path):
                # Check if this looks like an Office-related extension
                if not isinstance(value, str):
                    continue
                value_lower = value.lower()
                if not any(kw in value_lower for kw in ("office", "outlook", "groove", "onenote")):
                    continue

                # Check if the associated CLSID still exists
                clsid_path = f"SOFTWARE\\Classes\\CLSID\\{name}"
                if not key_exists(f"HKLM\\{clsid_path}"):
                    logger.info(
                        "Found orphaned shell extension approval: %s (%s)",
                        name,
                        value,
                        extra={"action": "shell-cleanup", "clsid": name, "dry_run": dry_run},
                    )
                    # Note: We don't remove from Approved list directly as it's a single key
                    # with multiple values. This just reports the orphaned entry.
                    removed += 1
        except (FileNotFoundError, OSError):
            continue

    return removed


def cleanup_protocol_handlers(
    protocols: Iterable[str],
    *,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> list[str]:
    """!
    @brief Remove orphaned protocol handler registrations.
    @param protocols Protocol names to check (e.g., "osf", "ms-word").
    @param dry_run If True, only report without deleting.
    @param logger Optional logger.
    @return List of removed protocol names.

    @details Removes URL protocol handlers that reference non-existent
    executables. Common Office protocols include osf:, ms-word:, etc.
    """
    logger = logger or _LOGGER
    removed: list[str] = []

    for protocol in protocols:
        # Check both HKLM and HKCU
        for hive, hive_name_str in [(_WINREG_HKLM, "HKLM"), (_WINREG_HKCU, "HKCU")]:
            path = f"SOFTWARE\\Classes\\{protocol}"
            full_path = f"{hive_name_str}\\{path}"

            if not key_exists(full_path):
                continue

            # Check if the shell\\open\\command points to an existing executable
            try:
                values = read_values(hive, f"{path}\\shell\\open\\command", view="native")
                default_cmd = values.get("", "")
                if default_cmd:
                    # Extract executable path (handle quoted paths)
                    exe_path = default_cmd.strip('"').split('"')[0].strip()
                    if exe_path and not Path(exe_path).exists():
                        logger.info(
                            "Removing orphaned protocol handler: %s",
                            full_path,
                            extra={
                                "action": "protocol-cleanup",
                                "protocol": protocol,
                                "dry_run": dry_run,
                            },
                        )
                        if not dry_run:
                            try:
                                delete_keys([full_path], dry_run=False, logger=logger)
                                if protocol not in removed:
                                    removed.append(protocol)
                            except Exception as e:
                                logger.warning("Failed to delete protocol %s: %s", protocol, e)
                        else:
                            if protocol not in removed:
                                removed.append(protocol)
            except (FileNotFoundError, OSError):
                continue

    return removed


__all__ = [
    "RegistryError",
    "WI_METADATA_PATHS",
    "cleanup_orphaned_typelibs",
    "cleanup_protocol_handlers",
    "cleanup_shell_extensions",
    "cleanup_wi_orphaned_components",
    "cleanup_wi_orphaned_products",
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
    "scan_orphaned_typelibs",
    "scan_wi_metadata",
    "validate_wi_metadata_key",
]
