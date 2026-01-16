"""!
@brief Office-specific registry detection and filtering utilities.
@details Provides heuristics and functions for identifying Office-related
registry entries, including GUID validation, uninstall entry detection,
and published components cleanup.
"""

from __future__ import annotations

import logging
from collections.abc import Callable, Iterable, Iterator, Mapping
from typing import Any

from . import registry_tools

_LOGGER = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Office Detection Constants
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# Office GUID Detection
# ---------------------------------------------------------------------------


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
            subkeys = list(registry_tools.iter_subkeys(hive, base_path, view=view))
        except FileNotFoundError:
            continue

        for subkey in subkeys:
            relative_path = f"{base_path}\\{subkey}"
            if (hive, relative_path) in seen_paths:
                continue
            values = registry_tools.read_values(hive, relative_path, view=view)
            if not values:
                continue
            if looks_like_office_entry(values):
                seen_paths.add((hive, relative_path))
                yield hive, relative_path, values


# ---------------------------------------------------------------------------
# Squished GUID Utilities
# ---------------------------------------------------------------------------


def decode_squished_guid(squished: str) -> str | None:
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


# ---------------------------------------------------------------------------
# Published Components Cleanup
# ---------------------------------------------------------------------------


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

    results: dict[str, Any] = {
        "entries_removed": [],
        "entries_kept": [],
        "value_deleted": False,
        "error": None,
    }

    try:
        # Read current value
        current_value = registry_tools.get_value(root, path, value_name, view=view)
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
        winreg = registry_tools.winreg
        if new_entries:
            # Write back the filtered list
            for mask in registry_tools._iter_access_masks(
                getattr(winreg, "KEY_WRITE", 0x20006), view
            ):
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
                for mask in registry_tools._iter_access_masks(
                    getattr(winreg, "KEY_WRITE", 0x20006), view
                ):
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

    results: dict[str, Any] = {
        "components_processed": 0,
        "values_modified": 0,
        "entries_removed": 0,
        "errors": [],
    }

    components_path = "Installer\\Components"
    hkcr = registry_tools._WINREG_HKCR

    def should_keep_entry(entry: str) -> bool:
        """Check if an entry references a non-Office product."""
        if len(entry) < 20:
            return True  # Too short to contain a valid squished GUID
        # First 20 chars are the squished GUID - decode and check
        squished = entry[:20]
        decoded = decode_squished_guid(squished)
        if decoded and is_office_guid(decoded):
            return False  # This is an Office entry, remove it
        return True  # Keep non-Office entries

    try:
        component_keys = list(registry_tools.iter_subkeys(hkcr, components_path, view=view))
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
            values = list(registry_tools.iter_values(hkcr, key_path, view=view))
        except (FileNotFoundError, OSError):
            continue

        for value_name, _value_type in values:
            result = filter_multi_string_value(
                hkcr,
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


__all__ = [
    "cleanup_published_components",
    "decode_squished_guid",
    "filter_multi_string_value",
    "is_office_guid",
    "iter_office_uninstall_entries",
    "looks_like_office_entry",
]
