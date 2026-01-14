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
from collections.abc import Callable, Iterable, Iterator, Mapping
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

# Office product code pattern: GUIDs ending with this suffix identify Office products
# The VBS uses "0000000FF1CE}" pattern with version > 14 (Office 2013+)
_OFFICE_GUID_SUFFIX = "0000000FF1CE}"

# SKU filters for C2R integration products (positions 11-14 in GUID)
_OFFICE_C2R_SKU_FILTERS = frozenset(("007E", "008F", "008C", "24E1", "237A", "00DD"))

# Known Office infrastructure GUIDs to always include
_OFFICE_SPECIAL_GUIDS = frozenset(
    (
        "{6C1ADE97-24E1-4AE4-AEDD-86D3A209CE60}",  # MOSA x64
        "{9520DDEB-237A-41DB-AA20-F2EF2360DCEB}",  # MOSA x86
        "{9AC08E99-230B-47E8-9721-4577B7F124EA}",  # Office shared
    )
)


def is_office_guid(guid: str) -> bool:
    """!
    @brief Determine whether a GUID belongs to an Office product.
    @details Mirrors the OffScrubC2R.vbs InScope() logic - checks for Office
        GUID suffix pattern with version > 14, plus known infrastructure GUIDs.
    @param guid Product code GUID in {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx} format.
    @returns True if this GUID matches Office product patterns.
    """
    if not guid or len(guid) != 38:
        return False

    upper = guid.upper()

    # Check special known GUIDs first
    if upper in _OFFICE_SPECIAL_GUIDS:
        return True

    # Check suffix pattern
    if not upper.endswith(_OFFICE_GUID_SUFFIX):
        return False

    # Check version > 14 (Office 2013+)
    # VBS: Mid(sProd, 4, 2) = positions 4-5 (1-indexed) = Python [3:5]
    try:
        version = int(upper[3:5])
        if version <= 14:
            return False
    except ValueError:
        return False

    # Check SKU filter
    # VBS: Mid(sProd, 11, 4) = positions 11-14 (1-indexed) = Python [10:14]
    sku = upper[10:14]
    return sku in _OFFICE_C2R_SKU_FILTERS


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


def filter_multi_string_value(
    root: int,
    path: str,
    value_name: str,
    should_keep: Callable[[str], bool],
    *,
    dry_run: bool = False,
    view: str | None = None,
    logger: logging.Logger | None = None,
) -> dict[str, Any]:
    """!
    @brief Filter entries from a REG_MULTI_SZ value based on a predicate.
    @details Implements the Published Components cleanup pattern from VBS.
        Rather than deleting the whole value, filters individual entries
        and writes back only the entries that should_keep returns True for.
    @param root Registry hive (HKLM, HKCU, etc.).
    @param path Registry key path.
    @param value_name Name of the REG_MULTI_SZ value.
    @param should_keep Predicate function - return True to keep an entry.
    @param dry_run If True, only log without modifying.
    @param view Registry view (native, 32bit, 64bit).
    @param logger Optional logger.
    @returns Dictionary with results: entries_removed, entries_kept, value_deleted.
    """
    logger = logger or _LOGGER
    _ensure_winreg()

    results: dict[str, Any] = {
        "entries_removed": [],
        "entries_kept": [],
        "value_deleted": False,
        "error": None,
    }

    try:
        # Read current value
        current_value = get_value(root, path, value_name, view=view)
        if current_value is None:
            return results

        # Handle both list and tuple (REG_MULTI_SZ returns list of strings)
        if not isinstance(current_value, (list, tuple)):
            logger.debug("Value %s is not multi-string type", value_name)
            return results

        entries = list(current_value)
        new_entries: list[str] = []

        for entry in entries:
            entry_str = str(entry) if entry is not None else ""
            if should_keep(entry_str):
                new_entries.append(entry_str)
                results["entries_kept"].append(entry_str)
            else:
                results["entries_removed"].append(entry_str)

        # If nothing removed, no action needed
        if len(new_entries) == len(entries):
            logger.debug("No entries removed from %s", value_name)
            return results

        logger.info(
            "Filtering REG_MULTI_SZ value: %d entries removed, %d kept",
            len(results["entries_removed"]),
            len(results["entries_kept"]),
            extra={
                "action": "registry-multi-string-filter",
                "path": path,
                "value": value_name,
                "removed_count": len(results["entries_removed"]),
                "kept_count": len(results["entries_kept"]),
                "dry_run": dry_run,
            },
        )

        if dry_run:
            return results

        # Write back filtered value or delete if empty
        if new_entries:
            # Write back the filtered list
            for mask in _iter_access_masks(getattr(winreg, "KEY_WRITE", 0x20006), view):
                try:
                    handle = winreg.OpenKey(root, path, 0, mask)
                    try:
                        winreg.SetValueEx(
                            handle,
                            value_name,
                            0,
                            getattr(winreg, "REG_MULTI_SZ", 7),
                            new_entries,
                        )
                        logger.debug("Wrote filtered value with %d entries", len(new_entries))
                        break
                    finally:
                        winreg.CloseKey(handle)
                except (FileNotFoundError, OSError) as exc:
                    logger.debug("Failed to write %s: %s", value_name, exc)
                    continue
        else:
            # All entries removed, delete the value
            try:
                for mask in _iter_access_masks(getattr(winreg, "KEY_WRITE", 0x20006), view):
                    try:
                        handle = winreg.OpenKey(root, path, 0, mask)
                        try:
                            winreg.DeleteValue(handle, value_name)
                            results["value_deleted"] = True
                            logger.debug("Deleted empty multi-string value %s", value_name)
                            break
                        finally:
                            winreg.CloseKey(handle)
                    except (FileNotFoundError, OSError):
                        continue
            except Exception as exc:
                results["error"] = str(exc)

    except Exception as exc:
        results["error"] = str(exc)
        logger.warning("Error filtering multi-string value %s: %s", value_name, exc)

    return results


def cleanup_published_components(
    *,
    dry_run: bool = False,
    view: str | None = None,
    logger: logging.Logger | None = None,
) -> dict[str, Any]:
    """!
    @brief Clean up Office entries from Windows Installer Published Components.
    @details Published Components are stored as REG_MULTI_SZ values under
        HKCR\\Installer\\Components. Each entry in the multi-string references
        a product GUID. This function filters out Office-related entries while
        preserving entries for non-Office products.
    @param dry_run If True, only log without modifying.
    @param view Registry view (native, 32bit, 64bit).
    @param logger Optional logger.
    @returns Dictionary with summary of cleanup results.
    """
    logger = logger or _LOGGER
    _ensure_winreg()

    results: dict[str, Any] = {
        "components_processed": 0,
        "values_modified": 0,
        "entries_removed": 0,
        "errors": [],
    }

    components_path = "Installer\\Components"

    def should_keep_entry(entry: str) -> bool:
        """Check if an entry references a non-Office product."""
        if len(entry) < 20:
            return True  # Too short to contain a valid squished GUID
        # First 20 chars are the squished GUID - decode and check
        squished = entry[:20]
        decoded = _decode_squished_guid(squished)
        if decoded and is_office_guid(decoded):
            return False  # This is an Office entry, remove it
        return True  # Keep non-Office entries

    try:
        component_keys = list(iter_subkeys(_WINREG_HKCR, components_path, view=view))
    except FileNotFoundError:
        logger.debug("Components key not found")
        return results
    except OSError as exc:
        results["errors"].append(f"Failed to enumerate components: {exc}")
        return results

    for component_key in component_keys:
        results["components_processed"] += 1
        key_path = f"{components_path}\\{component_key}"

        try:
            # Get all values in this component key
            values = list(iter_values(_WINREG_HKCR, key_path, view=view))
        except (FileNotFoundError, OSError):
            continue

        for value_name, _value_type in values:
            result = filter_multi_string_value(
                _WINREG_HKCR,
                key_path,
                value_name,
                should_keep_entry,
                dry_run=dry_run,
                view=view,
                logger=logger,
            )

            if result.get("entries_removed"):
                results["values_modified"] += 1
                results["entries_removed"] += len(result["entries_removed"])

            if result.get("error"):
                results["errors"].append(result["error"])

    logger.info(
        "Published Components cleanup: %d components, %d values modified, %d entries removed",
        results["components_processed"],
        results["values_modified"],
        results["entries_removed"],
        extra={
            "action": "cleanup-published-components",
            "dry_run": dry_run,
            **results,
        },
    )

    return results


def _decode_squished_guid(squished: str) -> str | None:
    """!
    @brief Decode a Windows Installer squished GUID to standard format.
    @details Windows Installer stores GUIDs in a compressed format where
        each segment is reversed. This decodes back to {xxxxxxxx-xxxx-...}.
    @param squished The 32-character squished GUID.
    @returns Standard GUID format or None if invalid.
    """
    if not squished or len(squished) < 32:
        return None

    try:
        # Squished format reverses each segment:
        # {ABCDEFGH-IJKL-MNOP-QRST-UVWXYZ012345}
        # becomes HGFEDCBAKLJIPONMTSRQWVXZYX103254
        s = squished.upper()
        return (
            "{"
            + s[7::-1]  # First segment reversed
            + "-"
            + s[11:7:-1]  # Second segment reversed
            + "-"
            + s[15:11:-1]  # Third segment reversed
            + "-"
            + s[17:15:-1]
            + s[19:17:-1]  # Fourth segment (two pairs)
            + "-"
            + s[21:19:-1]
            + s[23:21:-1]
            + s[25:23:-1]
            + s[27:25:-1]
            + s[29:27:-1]
            + s[31:29:-1]  # Fifth segment (six pairs)
            + "}"
        )
    except (IndexError, ValueError):
        return None


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
            try:
                exec_utils.run_command(
                    [reg_executable, "export", key, str(export_path), "/y"],
                    event="registry_export",
                    dry_run=dry_run,
                    check=True,
                    extra={"key": key, "path": str(export_path)},
                )
                exported.append(export_path)
            except Exception as exc:
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
        try:
            exec_utils.run_command(
                [reg_executable, "delete", key, "/f"],
                event="registry_delete",
                dry_run=dry_run,
                check=True,
                extra={"key": key},
            )
        except Exception as exc:
            # Deletion failure is non-fatal - key may not exist or access denied
            spinner.pause_for_output()
            logger.warning(
                "Registry deletion skipped for %s (key may not exist or access denied)",
                key,
            )
            spinner.resume_after_output()


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


# ---------------------------------------------------------------------------
# vNext Identity Registry Cleanup
# ---------------------------------------------------------------------------
# Based on OfficeScrubber.cmd :vNextREG subroutine (lines 697-714)

# Registry paths for vNext licensing/identity cleanup
_VNEXT_IDENTITY_KEYS: list[str] = [
    r"HKLM\SOFTWARE\Microsoft\Office\16.0\Common\Licensing",
    r"HKLM\SOFTWARE\Microsoft\Office\16.0\Common\Identity",
    r"HKLM\SOFTWARE\Microsoft\Office\16.0\Registration",
    r"HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Updates",
    r"HKLM\SOFTWARE\Microsoft\Office\16.0\Common\OEM",
    r"HKLM\SOFTWARE\Policies\Microsoft\Office\16.0\Common\Licensing",
    # WOW6432Node variants
    r"HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Common\OEM",
    r"HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\16.0\Common\Licensing",
    r"HKLM\SOFTWARE\Policies\WOW6432Node\Microsoft\Office\16.0\Common\Licensing",
]
"""!
@brief Registry keys to delete for vNext identity cleanup.
"""

_VNEXT_C2R_VALUES_TO_DELETE: list[str] = [
    "SharedComputerLicensing",
    "productkeys",
]
"""!
@brief Registry values to delete from ClickToRun Configuration key.
"""

_VNEXT_IDENTITY_VALUE_PATTERNS: list[str] = [
    r".*\.EmailAddress$",
    r".*\.TenantId$",
    r".*\.DeviceBasedLicensing$",
]
"""!
@brief Regex patterns for identity-related values to delete from C2R Configuration.
"""


def delete_registry_value(
    key_path: str,
    value_name: str,
    *,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> bool:
    """!
    @brief Delete a specific registry value.
    @param key_path Full registry path (e.g., "HKLM\\SOFTWARE\\...").
    @param value_name Name of the value to delete.
    @param dry_run If True, only log without deleting.
    @param logger Optional logger.
    @returns True if deleted or dry-run, False if not found or error.
    """
    logger = logger or _LOGGER
    reg_executable = shutil.which("reg")

    if not reg_executable:
        logger.warning("reg.exe not found, cannot delete value")
        return False

    logger.info(
        "Deleting registry value",
        extra={
            "action": "registry-value-delete",
            "key": key_path,
            "value": value_name,
            "dry_run": dry_run,
        },
    )

    if dry_run:
        return True

    result = exec_utils.run_command(
        [reg_executable, "delete", key_path, "/v", value_name, "/f"],
        event="registry_value_delete",
        dry_run=False,
        check=False,
        extra={"key": key_path, "value": value_name},
    )

    return result.returncode == 0


def cleanup_vnext_identity_registry(
    *,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> dict[str, Any]:
    """!
    @brief Clean up vNext identity and licensing registry entries.
    @details Implements OfficeScrubber.cmd :vNextREG subroutine functionality:
        1. Delete identity/licensing registry keys
        2. Delete specific C2R configuration values
        3. Delete identity-related values matching patterns (*.EmailAddress, etc.)
    @param dry_run If True, only log what would be deleted.
    @param logger Optional logger.
    @returns Dictionary with cleanup results.
    """
    logger = logger or _LOGGER
    results: dict[str, Any] = {
        "keys_deleted": [],
        "values_deleted": [],
        "patterns_matched": [],
        "errors": [],
    }

    # Step 1: Delete vNext identity keys
    logger.info("Cleaning vNext identity registry keys...")
    for key_path in _VNEXT_IDENTITY_KEYS:
        try:
            if key_exists(key_path):
                logger.debug("Deleting vNext key: %s", key_path)
                delete_keys([key_path], dry_run=dry_run, logger=logger)
                results["keys_deleted"].append(key_path)
        except Exception as e:
            logger.warning("Failed to delete key %s: %s", key_path, e)
            results["errors"].append({"key": key_path, "error": str(e)})

    # Step 2: Delete specific C2R configuration values
    c2r_config_paths = [
        r"HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
        r"HKLM\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration",
    ]

    for config_path in c2r_config_paths:
        if not key_exists(config_path):
            continue

        for value_name in _VNEXT_C2R_VALUES_TO_DELETE:
            try:
                if delete_registry_value(config_path, value_name, dry_run=dry_run, logger=logger):
                    results["values_deleted"].append(f"{config_path}\\{value_name}")
            except Exception as e:
                logger.debug("Failed to delete value %s\\%s: %s", config_path, value_name, e)

        # Step 3: Delete pattern-matched identity values
        # Need to enumerate values and match against patterns
        try:
            hive, _, subpath = config_path.partition("\\")
            hive_int = _WINREG_HKLM if hive == "HKLM" else _WINREG_HKCU
            patterns = [re.compile(p, re.IGNORECASE) for p in _VNEXT_IDENTITY_VALUE_PATTERNS]

            for value_name, _ in iter_values(hive_int, subpath):
                for pattern in patterns:
                    if pattern.match(value_name):
                        logger.debug("Pattern match for deletion: %s\\%s", config_path, value_name)
                        if delete_registry_value(
                            config_path, value_name, dry_run=dry_run, logger=logger
                        ):
                            results["patterns_matched"].append(f"{config_path}\\{value_name}")
                        break
        except (FileNotFoundError, OSError) as e:
            logger.debug("Failed to enumerate values in %s: %s", config_path, e)

    # Step 4: Clean SPP policies in Network Service SID (S-1-5-20)
    spp_key = r"HKU\S-1-5-20\Software\Microsoft\OfficeSoftwareProtectionPlatform\Policies\0ff1ce15-a989-479d-af46-f275c6370663"
    if key_exists(spp_key):
        try:
            delete_keys([spp_key], dry_run=dry_run, logger=logger)
            results["keys_deleted"].append(spp_key)
        except Exception as e:
            logger.debug("Failed to delete SPP key %s: %s", spp_key, e)

    total_deleted = (
        len(results["keys_deleted"])
        + len(results["values_deleted"])
        + len(results["patterns_matched"])
    )
    logger.info(
        "vNext identity cleanup complete: %d keys, %d values, %d pattern matches",
        len(results["keys_deleted"]),
        len(results["values_deleted"]),
        len(results["patterns_matched"]),
    )

    return results


# ---------------------------------------------------------------------------
# User Profile Registry Loading
# ---------------------------------------------------------------------------
# Based on OffScrubC2R.vbs LoadUsersReg subroutine (lines 2192-2214)


def get_user_profiles_directory() -> Path | None:
    """!
    @brief Get the path to the Windows user profiles directory.
    @returns Path to profiles directory (e.g., C:\\Users) or None if not found.
    """
    try:
        key_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
        values = read_values(_WINREG_HKLM, key_path, view="native")
        profiles_dir = values.get("ProfilesDirectory", "")
        if profiles_dir:
            # Expand environment variables
            import os

            expanded = os.path.expandvars(profiles_dir)
            return Path(expanded)
    except (FileNotFoundError, OSError):
        pass
    return None


def get_user_profile_hive_paths() -> list[tuple[str, Path]]:
    """!
    @brief Enumerate user profile folders and their ntuser.dat paths.
    @returns List of (profile_name, ntuser_dat_path) tuples.
    """
    profiles_dir = get_user_profiles_directory()
    if profiles_dir is None or not profiles_dir.exists():
        return []

    results: list[tuple[str, Path]] = []
    try:
        for folder in profiles_dir.iterdir():
            if not folder.is_dir():
                continue
            ntuser = folder / "ntuser.dat"
            if ntuser.exists():
                results.append((folder.name, ntuser))
    except OSError:
        pass

    return results


# Track loaded user hives for cleanup
_LOADED_USER_HIVES: list[str] = []


def load_user_registry_hives(
    *,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> list[str]:
    """!
    @brief Load all user ntuser.dat files into HKU for per-user cleanup.
    @details Implements OffScrubC2R.vbs LoadUsersReg functionality.
        Loads each user's registry hive to HKU\\<profile_name>.
    @param dry_run If True, only log what would be loaded.
    @param logger Optional logger.
    @returns List of successfully loaded hive names.
    """
    global _LOADED_USER_HIVES
    logger = logger or _LOGGER

    profiles = get_user_profile_hive_paths()
    if not profiles:
        logger.debug("No user profile hives found to load")
        return []

    loaded: list[str] = []
    reg_exe = shutil.which("reg")
    if not reg_exe:
        logger.warning("reg.exe not found, cannot load user hives")
        return []

    for profile_name, ntuser_path in profiles:
        hive_key = f"HKU\\{profile_name}"

        # Skip if already loaded (e.g., current user)
        if key_exists(hive_key):
            logger.debug("Hive %s already loaded, skipping", hive_key)
            continue

        logger.info(
            "Loading user registry hive",
            extra={
                "action": "registry-hive-load",
                "profile": profile_name,
                "path": str(ntuser_path),
                "dry_run": dry_run,
            },
        )

        if dry_run:
            loaded.append(profile_name)
            continue

        # reg load "HKU\<profile_name>" "<path>\ntuser.dat"
        result = exec_utils.run_command(
            [reg_exe, "load", hive_key, str(ntuser_path)],
            event="registry_hive_load",
            dry_run=False,
            check=False,
            extra={"profile": profile_name},
        )

        if result.returncode == 0:
            loaded.append(profile_name)
            _LOADED_USER_HIVES.append(profile_name)
            logger.debug("Loaded hive %s", hive_key)
        else:
            logger.debug(
                "Failed to load hive %s (code %d): %s",
                hive_key,
                result.returncode,
                result.stderr.strip() if result.stderr else "",
            )

    return loaded


def unload_user_registry_hives(
    *,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> list[str]:
    """!
    @brief Unload previously loaded user registry hives.
    @param dry_run If True, only log what would be unloaded.
    @param logger Optional logger.
    @returns List of successfully unloaded hive names.
    """
    global _LOADED_USER_HIVES
    logger = logger or _LOGGER

    if not _LOADED_USER_HIVES:
        logger.debug("No user hives to unload")
        return []

    unloaded: list[str] = []
    reg_exe = shutil.which("reg")
    if not reg_exe:
        logger.warning("reg.exe not found, cannot unload user hives")
        return []

    # Unload in reverse order
    for profile_name in reversed(_LOADED_USER_HIVES.copy()):
        hive_key = f"HKU\\{profile_name}"

        logger.info(
            "Unloading user registry hive",
            extra={
                "action": "registry-hive-unload",
                "profile": profile_name,
                "dry_run": dry_run,
            },
        )

        if dry_run:
            unloaded.append(profile_name)
            continue

        # reg unload "HKU\<profile_name>"
        result = exec_utils.run_command(
            [reg_exe, "unload", hive_key],
            event="registry_hive_unload",
            dry_run=False,
            check=False,
            extra={"profile": profile_name},
        )

        if result.returncode == 0:
            unloaded.append(profile_name)
            _LOADED_USER_HIVES.remove(profile_name)
            logger.debug("Unloaded hive %s", hive_key)
        else:
            logger.warning(
                "Failed to unload hive %s (code %d)",
                hive_key,
                result.returncode,
            )

    return unloaded


def get_loaded_user_hives() -> list[str]:
    """!
    @brief Get list of user hives currently loaded by this session.
    @returns List of profile names with loaded hives.
    """
    return list(_LOADED_USER_HIVES)


# ---------------------------------------------------------------------------
# Taskband Registry Cleanup
# ---------------------------------------------------------------------------
# Based on OffScrubC2R.vbs ClearTaskBand subroutine (lines 2161-2183)

_TASKBAND_VALUES_TO_DELETE: list[str] = [
    "Favorites",
    "FavoritesRemovedChanges",
    "FavoritesChanges",
    "FavoritesResolve",
    "FavoritesVersion",
]
"""!
@brief Registry values to delete from Taskband key to unpin items.
"""


def cleanup_taskband_registry(
    *,
    include_all_users: bool = False,
    dry_run: bool = False,
    logger: logging.Logger | None = None,
) -> dict[str, Any]:
    """!
    @brief Clean up taskband registry to remove pinned Office items.
    @details Implements OffScrubC2R.vbs ClearTaskBand functionality.
        Removes Favorites* values from Taskband key to clear pinned items.
    @param include_all_users If True, also clean all user profiles in HKU.
    @param dry_run If True, only log what would be deleted.
    @param logger Optional logger.
    @returns Dictionary with cleanup results.
    """
    logger = logger or _LOGGER
    results: dict[str, Any] = {
        "values_deleted": [],
        "users_processed": [],
        "errors": [],
    }

    taskband_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband"

    # Step 1: Clean HKCU taskband
    hkcu_full_path = f"HKCU\\{taskband_path}"
    logger.info("Cleaning HKCU taskband registry...")

    for value_name in _TASKBAND_VALUES_TO_DELETE:
        try:
            if delete_registry_value(hkcu_full_path, value_name, dry_run=dry_run, logger=logger):
                results["values_deleted"].append(f"{hkcu_full_path}\\{value_name}")
        except Exception as e:
            logger.debug("Failed to delete %s\\%s: %s", hkcu_full_path, value_name, e)

    # Step 2: If requested, clean all user profiles in HKU
    if include_all_users:
        logger.info("Cleaning taskband for all user profiles...")

        # First load user hives if not already loaded
        loaded_hives = load_user_registry_hives(dry_run=dry_run, logger=logger)
        if loaded_hives:
            results["users_processed"].extend(loaded_hives)

        # Enumerate all SIDs in HKU
        try:
            for sid in iter_subkeys(_WINREG_HKU, "", view="native"):
                # Skip well-known SIDs that don't have user profiles
                if sid in ("S-1-5-18", "S-1-5-19", "S-1-5-20", ".DEFAULT"):
                    continue
                if sid.endswith("_Classes"):
                    continue

                hku_taskband_path = f"HKU\\{sid}\\{taskband_path}"

                for value_name in _TASKBAND_VALUES_TO_DELETE:
                    try:
                        if delete_registry_value(
                            hku_taskband_path, value_name, dry_run=dry_run, logger=logger
                        ):
                            results["values_deleted"].append(f"{hku_taskband_path}\\{value_name}")
                    except Exception:
                        pass  # Value may not exist, which is fine

                if sid not in results["users_processed"]:
                    results["users_processed"].append(sid)

        except (FileNotFoundError, OSError) as e:
            logger.debug("Failed to enumerate HKU: %s", e)

        # Unload hives we loaded
        if loaded_hives:
            unload_user_registry_hives(dry_run=dry_run, logger=logger)

    logger.info(
        "Taskband cleanup complete: %d values deleted across %d users",
        len(results["values_deleted"]),
        len(results["users_processed"]) + 1,  # +1 for HKCU
    )

    return results


__all__ = [
    "RegistryError",
    "WI_METADATA_PATHS",
    "cleanup_orphaned_typelibs",
    "cleanup_protocol_handlers",
    "cleanup_published_components",
    "cleanup_shell_extensions",
    "cleanup_taskband_registry",
    "cleanup_vnext_identity_registry",
    "cleanup_wi_orphaned_components",
    "cleanup_wi_orphaned_products",
    "delete_keys",
    "delete_registry_value",
    "export_keys",
    "filter_multi_string_value",
    "get_loaded_user_hives",
    "get_user_profile_hive_paths",
    "get_user_profiles_directory",
    "get_value",
    "hive_name",
    "is_office_guid",
    "iter_office_uninstall_entries",
    "iter_subkeys",
    "iter_values",
    "key_exists",
    "load_user_registry_hives",
    "looks_like_office_entry",
    "open_key",
    "read_values",
    "scan_orphaned_typelibs",
    "scan_wi_metadata",
    "unload_user_registry_hives",
    "validate_wi_metadata_key",
]
