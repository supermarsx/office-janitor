"""!
@brief GUID manipulation utilities for Windows Installer operations.
@details Implements the GUID compression, expansion, and squishing algorithms
used by Windows Installer to store product and component identifiers in the
registry. These transformations are essential for navigating WI metadata
stored under ``HKLM\\SOFTWARE\\Classes\\Installer`` and user-specific paths.

The algorithms mirror the VBS implementations in the legacy OffScrub scripts:
- ``GetCompressedGuid`` / ``GetExpandedGuid`` from OffScrub_O16msi.vbs
- ``GetSquishGuid`` / ``GetDecodeSquishGuid`` from OffScrubC2R.vbs

@note Windows Installer uses three GUID formats:
1. **Standard**: ``{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`` (38 chars with braces)
2. **Compressed**: 32-char reversed-nibble format used in registry paths
3. **Squished**: 20-char format used in some WI component paths
"""

from __future__ import annotations

import re
from typing import Final

# Standard GUID pattern: {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}
_GUID_PATTERN: Final[re.Pattern[str]] = re.compile(
    r"^\{?([0-9A-Fa-f]{8})-?([0-9A-Fa-f]{4})-?([0-9A-Fa-f]{4})-?"
    r"([0-9A-Fa-f]{4})-?([0-9A-Fa-f]{12})\}?$"
)

# Compressed GUID pattern: 32 hex characters
_COMPRESSED_PATTERN: Final[re.Pattern[str]] = re.compile(r"^[0-9A-Fa-f]{32}$")

# Squished GUID pattern: 20 alphanumeric characters (base85-like encoding)
_SQUISHED_PATTERN: Final[re.Pattern[str]] = re.compile(r"^[0-9A-Za-z]{20}$")


class GuidError(ValueError):
    """!
    @brief Raised when GUID parsing or transformation fails.
    """


def _reverse_pairs(s: str) -> str:
    """!
    @brief Reverse each pair of characters in a string.
    @param s Input string (must have even length).
    @return String with character pairs reversed.

    @details This is the core operation used by Windows Installer to create
    compressed GUIDs. For example, "ABCD" becomes "BADC".
    """
    return "".join(s[i + 1] + s[i] for i in range(0, len(s), 2))


def is_valid_guid(guid: str) -> bool:
    """!
    @brief Check if a string is a valid standard GUID format.
    @param guid String to validate.
    @return True if valid GUID format, False otherwise.
    """
    return _GUID_PATTERN.match(guid) is not None


def is_compressed_guid(guid: str) -> bool:
    """!
    @brief Check if a string is a valid compressed GUID format.
    @param guid String to validate.
    @return True if valid compressed format (32 hex chars), False otherwise.
    """
    return _COMPRESSED_PATTERN.match(guid) is not None


def is_squished_guid(guid: str) -> bool:
    """!
    @brief Check if a string is a valid squished GUID format.
    @param guid String to validate.
    @return True if valid squished format (20 alphanumeric chars), False otherwise.
    """
    return _SQUISHED_PATTERN.match(guid) is not None


def compress_guid(guid: str) -> str:
    """!
    @brief Convert a standard GUID to Windows Installer compressed format.
    @param guid Standard GUID string (with or without braces/hyphens).
    @return 32-character compressed GUID.
    @throws GuidError If the input is not a valid GUID.

    @details The compression algorithm:
    1. Remove braces and hyphens
    2. Split into segments: 8-4-4-2-2-2-2-2-2-2-2 characters
    3. Reverse each pair of characters within each segment

    Example:
    - Input:  ``{00000000-0000-0000-0000-000000000001}``
    - Output: ``00000000000000000000000000000010``

    This matches the VBS ``GetCompressedGuid`` function from OffScrub_O16msi.vbs.
    """
    match = _GUID_PATTERN.match(guid)
    if not match:
        raise GuidError(f"Invalid GUID format: {guid}")

    # Extract groups: (8 chars)-(4)-(4)-(4)-(12)
    g1, g2, g3, g4, g5 = match.groups()

    # Apply pair reversal to each segment
    # The 4th group (4 chars) and 5th group (12 chars) are treated as pairs
    compressed = (
        _reverse_pairs(g1)
        + _reverse_pairs(g2)
        + _reverse_pairs(g3)
        + _reverse_pairs(g4)
        + _reverse_pairs(g5)
    )

    return compressed.upper()


def expand_guid(compressed: str) -> str:
    """!
    @brief Convert a compressed GUID back to standard format.
    @param compressed 32-character compressed GUID.
    @return Standard GUID with braces and hyphens.
    @throws GuidError If the input is not a valid compressed GUID.

    @details Reverses the compression algorithm to restore the original GUID.

    Example:
    - Input:  ``00000000000000000000000000000010``
    - Output: ``{00000000-0000-0000-0000-000000000001}``

    This matches the VBS ``GetExpandedGuid`` function from OffScrub_O16msi.vbs.
    """
    if not is_compressed_guid(compressed):
        raise GuidError(f"Invalid compressed GUID format: {compressed}")

    # Split into segments matching the compression pattern
    c = compressed.upper()
    g1 = _reverse_pairs(c[0:8])
    g2 = _reverse_pairs(c[8:12])
    g3 = _reverse_pairs(c[12:16])
    g4 = _reverse_pairs(c[16:20])
    g5 = _reverse_pairs(c[20:32])

    return f"{{{g1}-{g2}-{g3}-{g4}-{g5}}}"


def normalize_guid(guid: str) -> str:
    """!
    @brief Normalize a GUID to standard uppercase format with braces.
    @param guid GUID in any valid format (standard, compressed, or no braces).
    @return Standard GUID format: ``{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}``
    @throws GuidError If the input is not a valid GUID.
    """
    # If it's compressed, expand it first
    if is_compressed_guid(guid):
        return expand_guid(guid)

    # Parse and reformat standard GUID
    match = _GUID_PATTERN.match(guid)
    if not match:
        raise GuidError(f"Invalid GUID format: {guid}")

    g1, g2, g3, g4, g5 = match.groups()
    return f"{{{g1}-{g2}-{g3}-{g4}-{g5}}}".upper()


def strip_guid_braces(guid: str) -> str:
    """!
    @brief Remove braces from a GUID string.
    @param guid Standard GUID (with or without braces).
    @return GUID without braces: ``XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX``
    """
    return guid.strip("{}").upper()


def squish_guid(guid: str) -> str:
    """!
    @brief Convert a GUID to 20-character squished format.
    @param guid Standard GUID string.
    @return 20-character squished representation.
    @throws GuidError If the input is not a valid GUID.

    @details The squished format encodes the 128-bit GUID as a 20-character
    alphanumeric string using a base85-like encoding. This format appears in
    some Windows Installer component paths.

    The algorithm (from VBS ``GetSquishGuid``):
    1. Remove braces and hyphens to get 32 hex chars
    2. Process in groups of 8 hex chars (32 bits each)
    3. Encode each 32-bit value as 5 base-85 characters

    @note The exact encoding uses a custom character set, not standard base85.
    """
    match = _GUID_PATTERN.match(guid)
    if not match:
        raise GuidError(f"Invalid GUID format: {guid}")

    # Get raw 32 hex characters
    hex_str = "".join(match.groups())

    # Character set for squished encoding (custom base-85 alphabet)
    # This matches the VBS implementation's character mapping
    chars = (
        "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
        "!#$%&()*+,-./:;<=>?@[]^_`{|}~"
    )

    result: list[str] = []

    # Process 8 hex chars (32 bits) at a time, producing 5 output chars
    for i in range(0, 32, 8):
        chunk = hex_str[i : i + 8]
        value = int(chunk, 16)

        # Encode as 5 base-85 characters (85^5 > 2^32)
        encoded: list[str] = []
        for _ in range(5):
            encoded.append(chars[value % 85])
            value //= 85
        result.extend(encoded)

    return "".join(result)


def decode_squished_guid(squished: str) -> str:
    """!
    @brief Convert a squished GUID back to standard format.
    @param squished 20-character squished GUID.
    @return Standard GUID with braces and hyphens.
    @throws GuidError If the input is not a valid squished GUID.

    @details Reverses the squishing algorithm to restore the original GUID.

    This matches the VBS ``GetDecodeSquishGuid`` function from OffScrubC2R.vbs.
    """
    if len(squished) != 20:
        raise GuidError(f"Invalid squished GUID length: {len(squished)} (expected 20)")

    # Character set for squished encoding
    chars = (
        "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
        "!#$%&()*+,-./:;<=>?@[]^_`{|}~"
    )
    char_to_val = {c: i for i, c in enumerate(chars)}

    hex_parts: list[str] = []

    # Decode 5 chars at a time to 8 hex chars (32 bits)
    for i in range(0, 20, 5):
        chunk = squished[i : i + 5]

        # Validate characters
        for c in chunk:
            if c not in char_to_val:
                raise GuidError(f"Invalid character in squished GUID: {c}")

        # Decode base-85 to integer
        value = 0
        for j, c in enumerate(chunk):
            value += char_to_val[c] * (85**j)

        # Convert to 8 hex characters
        hex_parts.append(f"{value:08X}")

    # Reconstruct GUID from hex string
    hex_str = "".join(hex_parts)
    return f"{{{hex_str[0:8]}-{hex_str[8:12]}-{hex_str[12:16]}-{hex_str[16:20]}-{hex_str[20:32]}}}"


def guid_to_registry_path(guid: str, base_path: str) -> str:
    """!
    @brief Build a WI registry path for a product or component GUID.
    @param guid Standard GUID string.
    @param base_path Base registry path (e.g., ``SOFTWARE\\Classes\\Installer\\Products``).
    @return Full registry path with compressed GUID appended.

    @details Windows Installer stores product and component metadata under
    paths like:
    - ``HKLM\\SOFTWARE\\Classes\\Installer\\Products\\<compressed_guid>``
    - ``HKLM\\SOFTWARE\\Classes\\Installer\\Components\\<compressed_guid>``
    """
    compressed = compress_guid(guid)
    return f"{base_path}\\{compressed}"


def extract_guid_from_path(path: str) -> str | None:
    """!
    @brief Extract and expand a compressed GUID from a registry path.
    @param path Registry path potentially containing a compressed GUID.
    @return Expanded GUID if found, None otherwise.

    @details Searches for 32-character hex sequences in the path and attempts
    to expand them as compressed GUIDs.
    """
    # Look for 32 consecutive hex characters
    match = re.search(r"[0-9A-Fa-f]{32}", path)
    if match:
        try:
            return expand_guid(match.group())
        except GuidError:
            return None
    return None


# Office-specific product code patterns (from VBS scripts)
# Product codes follow the pattern: {XXXXXXXX-XXXX-XXXX-YYYY-ZZZZZZZZZZZZ}
# where YYYY indicates the product type

OFFICE_PRODUCT_TYPE_CODES: dict[str, str] = {
    "0000": "Unknown/Other",
    "000F": "Professional Plus (Volume)",
    "0011": "Professional Plus (Retail)",
    "0012": "Standard",
    "0013": "Home and Business",
    "0014": "Home and Student",
    "0015": "Access",
    "0016": "Excel",
    "0017": "SharePoint Designer",
    "0018": "PowerPoint",
    "0019": "Publisher",
    "001A": "Outlook",
    "001B": "Word",
    "001C": "Access Runtime",
    "001F": "Proofing Tools",
    "002E": "Ultimate",
    "002F": "Home and Student (Retail)",
    "003A": "Project Standard",
    "003B": "Project Professional",
    "0044": "InfoPath",
    "0051": "Visio Professional",
    "0052": "Visio Premium",
    "0053": "Visio Standard",
    "0057": "Visio",
    "00A1": "OneNote",
    "00A3": "OneNote (Retail)",
    "00A7": "Calendar Printing Assistant",
    "00A9": "InterConnect",
    "00AF": "PowerPoint Viewer",
    "00B0": "Save as PDF",
    "00B1": "Save as XPS",
    "00B2": "Save as PDF/XPS",
    "00BA": "Groove",
    "00CA": "Small Business Basics",
    "00E0": "Outlook Connector",
    "00FD": "Lync Basic",
    "012B": "Lync",
    "012C": "Lync (Retail)",
    "0131": "Lync Trial",
    "0135": "Lync Basic",
}
"""!
@brief Mapping of Office product type codes to human-readable names.
@details The product type code is extracted from positions 21-24 of the GUID
(after removing braces and hyphens). This matches the VBS ``GetProductType``
logic from OffScrub_O16msi.vbs.
"""


def get_product_type_code(product_code: str) -> str | None:
    """!
    @brief Extract the product type code from an Office product GUID.
    @param product_code Standard Office product GUID.
    @return 4-character product type code, or None if invalid.

    @details Office product GUIDs encode the product type in a specific position.
    For example, ``{90160000-000F-0000-1000-0000000FF1CE}`` has type ``000F``
    (Professional Plus Volume).
    """
    match = _GUID_PATTERN.match(product_code)
    if not match:
        return None

    # Type code is in the second group (positions 9-12 after the first hyphen)
    g2 = match.group(2).upper()
    return g2


def classify_office_product(product_code: str) -> str:
    """!
    @brief Classify an Office product by its type code.
    @param product_code Standard Office product GUID.
    @return Human-readable product classification.
    """
    type_code = get_product_type_code(product_code)
    if type_code is None:
        return "Unknown"
    return OFFICE_PRODUCT_TYPE_CODES.get(type_code, f"Unknown ({type_code})")


def is_office_product_code(product_code: str) -> bool:
    """!
    @brief Check if a product code appears to be an Office product.
    @param product_code Standard product GUID.
    @return True if the GUID matches Office product patterns.

    @details Office products typically end with ``0000000FF1CE`` (O-F-F-I-C-E).
    """
    match = _GUID_PATTERN.match(product_code)
    if not match:
        return False

    # Check if the last 12 characters match the Office signature
    g5 = match.group(5).upper()
    return g5.endswith("FF1CE") or g5.endswith("F1CE0")


# Alias for backward compatibility
is_office_guid = is_office_product_code
"""!
@brief Alias for is_office_product_code.
"""


def get_office_version_from_guid(product_code: str) -> str | None:
    """!
    @brief Extract the Office major version from a product GUID.
    @param product_code Standard product GUID.
    @return Version string (e.g., "15", "16") or None if not an Office product.

    @details Office product codes encode the version in positions 5-6 (after the brace):
        - 90 = Office 2003
        - 91 = Office 2007
        - 92 = Office 2010 (Volume License)
        - 93 = Office 2010 (Retail)
        - 94 = Office 2013 (Click-to-Run)
        - 95 = Office 2013 (MSI)
        - A1 = Office 2016/365 (Click-to-Run)
        - 90140 = Office 2010
        - 15.0 = Office 2013
        - 16.0 = Office 2016+
    """
    if not is_office_product_code(product_code):
        return None

    match = _GUID_PATTERN.match(product_code)
    if not match:
        return None

    # Extract first group and get version indicator
    g1 = match.group(1).upper()

    # Office 2016+ (16.0) typically starts with 9
    # but we look at specific patterns
    version_map = {
        "90": "11",  # Office 2003
        "91": "12",  # Office 2007
        "92": "14",  # Office 2010 VL
        "93": "14",  # Office 2010 Retail
        "94": "15",  # Office 2013 C2R
        "95": "15",  # Office 2013 MSI
        "A1": "16",  # Office 2016+ C2R
    }

    # Check first 2 chars of first group
    prefix = g1[:2]
    if prefix in version_map:
        return version_map[prefix]

    # Try to infer from other patterns
    # Most modern Office uses patterns starting with 9
    if g1.startswith("9"):
        return "16"

    return None


__all__ = [
    "GuidError",
    "classify_office_product",
    "compress_guid",
    "decode_squished_guid",
    "expand_guid",
    "extract_guid_from_path",
    "get_office_version_from_guid",
    "get_product_type_code",
    "guid_to_registry_path",
    "is_compressed_guid",
    "is_office_guid",
    "is_office_product_code",
    "is_squished_guid",
    "is_valid_guid",
    "normalize_guid",
    "squish_guid",
    "strip_guid_braces",
    "OFFICE_PRODUCT_TYPE_CODES",
]
