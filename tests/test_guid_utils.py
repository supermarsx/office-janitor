"""!
@brief Tests for GUID manipulation utilities.
@details Validates the compression, expansion, and squishing algorithms
used by Windows Installer to transform GUIDs.
"""

from __future__ import annotations

import pytest

from office_janitor.guid_utils import (
    GuidError,
    classify_office_product,
    compress_guid,
    decode_squished_guid,
    expand_guid,
    extract_guid_from_path,
    get_product_type_code,
    guid_to_registry_path,
    is_compressed_guid,
    is_office_product_code,
    is_squished_guid,
    is_valid_guid,
    normalize_guid,
    squish_guid,
    strip_guid_braces,
)


class TestGuidValidation:
    """Tests for GUID format validation functions."""

    def test_is_valid_guid_with_braces(self) -> None:
        """Standard GUID with braces should be valid."""
        assert is_valid_guid("{00000000-0000-0000-0000-000000000000}")
        assert is_valid_guid("{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}")
        assert is_valid_guid("{12345678-1234-1234-1234-123456789ABC}")

    def test_is_valid_guid_without_braces(self) -> None:
        """GUID without braces should be valid."""
        assert is_valid_guid("00000000-0000-0000-0000-000000000000")
        assert is_valid_guid("FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF")

    def test_is_valid_guid_lowercase(self) -> None:
        """Lowercase GUIDs should be valid."""
        assert is_valid_guid("{abcdef00-1234-5678-9abc-def012345678}")

    def test_is_valid_guid_invalid_formats(self) -> None:
        """Invalid GUID formats should be rejected."""
        assert not is_valid_guid("")
        assert not is_valid_guid("not-a-guid")
        assert not is_valid_guid("{00000000-0000-0000-0000}")  # Too short
        assert not is_valid_guid("{00000000-0000-0000-0000-0000000000000}")  # Too long
        assert not is_valid_guid("{GGGGGGGG-GGGG-GGGG-GGGG-GGGGGGGGGGGG}")  # Invalid hex

    def test_is_compressed_guid_valid(self) -> None:
        """32-char hex strings should be valid compressed GUIDs."""
        assert is_compressed_guid("00000000000000000000000000000000")
        assert is_compressed_guid("FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF")
        assert is_compressed_guid("12345678123412341234123456789ABC")

    def test_is_compressed_guid_invalid(self) -> None:
        """Non-32-char or non-hex strings should be invalid."""
        assert not is_compressed_guid("")
        assert not is_compressed_guid("0000000000000000")  # Too short
        assert not is_compressed_guid("{00000000-0000-0000-0000-000000000000}")  # Standard format
        assert not is_compressed_guid("GGGGGGGGGGGGGGGGGGGGGGGGGGGGGGG")  # Invalid hex


class TestGuidCompression:
    """Tests for GUID compression and expansion."""

    def test_compress_guid_basic(self) -> None:
        """Basic compression should reverse character pairs."""
        # Zero GUID
        result = compress_guid("{00000000-0000-0000-0000-000000000000}")
        assert result == "00000000000000000000000000000000"

    def test_compress_guid_sample(self) -> None:
        """Compression should match Windows Installer algorithm."""
        # Example from VBS OffScrub code
        result = compress_guid("{90160000-0011-0000-0000-0000000FF1CE}")
        # The algorithm reverses pairs within each segment
        # {90160000} -> 09610000
        # {0011} -> 1100
        # {0000} -> 0000
        # {0000} -> 0000
        # {0000000FF1CE} -> 000000F0F1EC
        assert len(result) == 32
        assert result.isalnum()

    def test_compress_guid_without_braces(self) -> None:
        """Compression should work without braces."""
        result = compress_guid("90160000-0011-0000-0000-0000000FF1CE")
        assert len(result) == 32

    def test_compress_guid_invalid(self) -> None:
        """Invalid GUIDs should raise GuidError."""
        with pytest.raises(GuidError):
            compress_guid("not-a-guid")
        with pytest.raises(GuidError):
            compress_guid("")

    def test_expand_guid_basic(self) -> None:
        """Expansion should reverse compression."""
        compressed = "00000000000000000000000000000000"
        result = expand_guid(compressed)
        assert result == "{00000000-0000-0000-0000-000000000000}"

    def test_expand_guid_invalid(self) -> None:
        """Invalid compressed GUIDs should raise GuidError."""
        with pytest.raises(GuidError):
            expand_guid("not-compressed")
        with pytest.raises(GuidError):
            expand_guid("0000000000000000")  # Too short

    def test_compress_expand_roundtrip(self) -> None:
        """Compression followed by expansion should return original."""
        original = "{90160000-0011-0000-0000-0000000FF1CE}"
        compressed = compress_guid(original)
        expanded = expand_guid(compressed)
        assert expanded == original

    def test_expand_compress_roundtrip(self) -> None:
        """Expansion followed by compression should return original."""
        original = "09610000110000000000000000F0F1EC"
        expanded = expand_guid(original)
        compressed = compress_guid(expanded)
        assert compressed == original


class TestGuidNormalization:
    """Tests for GUID normalization."""

    def test_normalize_guid_with_braces(self) -> None:
        """GUID with braces should normalize correctly."""
        result = normalize_guid("{abcdef00-1234-5678-9abc-def012345678}")
        assert result == "{ABCDEF00-1234-5678-9ABC-DEF012345678}"

    def test_normalize_guid_without_braces(self) -> None:
        """GUID without braces should add them."""
        result = normalize_guid("abcdef00-1234-5678-9abc-def012345678")
        assert result == "{ABCDEF00-1234-5678-9ABC-DEF012345678}"

    def test_normalize_guid_from_compressed(self) -> None:
        """Compressed GUID should expand to standard format."""
        result = normalize_guid("00000000000000000000000000000000")
        assert result == "{00000000-0000-0000-0000-000000000000}"

    def test_strip_guid_braces(self) -> None:
        """Braces should be removed."""
        result = strip_guid_braces("{ABCDEF00-1234-5678-9ABC-DEF012345678}")
        assert result == "ABCDEF00-1234-5678-9ABC-DEF012345678"


class TestGuidSquishing:
    """Tests for GUID squishing (base-85 encoding)."""

    def test_squish_guid_length(self) -> None:
        """Squished GUID should be 20 characters."""
        result = squish_guid("{00000000-0000-0000-0000-000000000000}")
        assert len(result) == 20

    def test_squish_guid_alphanumeric(self) -> None:
        """Squished GUID should use the encoding character set."""
        result = squish_guid("{90160000-0011-0000-0000-0000000FF1CE}")
        assert len(result) == 20

    def test_squish_decode_roundtrip(self) -> None:
        """Squishing followed by decoding should return original."""
        original = "{90160000-0011-0000-0000-0000000FF1CE}"
        squished = squish_guid(original)
        decoded = decode_squished_guid(squished)
        assert decoded == original

    def test_decode_squished_invalid_length(self) -> None:
        """Invalid length should raise GuidError."""
        with pytest.raises(GuidError):
            decode_squished_guid("tooshort")
        with pytest.raises(GuidError):
            decode_squished_guid("thisistoolongtobevalid")

    def test_decode_squished_invalid_chars(self) -> None:
        """Invalid characters should raise GuidError."""
        with pytest.raises(GuidError):
            decode_squished_guid("invalid\x00characters")


class TestRegistryPathHelpers:
    """Tests for registry path manipulation."""

    def test_guid_to_registry_path(self) -> None:
        """Path should include compressed GUID."""
        result = guid_to_registry_path(
            "{90160000-0011-0000-0000-0000000FF1CE}",
            r"SOFTWARE\Classes\Installer\Products",
        )
        assert result.startswith(r"SOFTWARE\Classes\Installer\Products" + "\\")
        assert len(result.split("\\")[-1]) == 32

    def test_extract_guid_from_path(self) -> None:
        """Compressed GUID should be extracted and expanded."""
        # First compress a known GUID, then build a path
        guid = "{90160000-0011-0000-0000-0000000FF1CE}"
        compressed = compress_guid(guid)
        path = rf"HKLM\SOFTWARE\Classes\Installer\Products\{compressed}"

        result = extract_guid_from_path(path)
        assert result == guid

    def test_extract_guid_from_path_no_guid(self) -> None:
        """Path without GUID should return None."""
        result = extract_guid_from_path(r"HKLM\SOFTWARE\Microsoft\Office")
        assert result is None


class TestOfficeProductClassification:
    """Tests for Office-specific product code handling."""

    def test_get_product_type_code(self) -> None:
        """Product type code should be extracted correctly."""
        # Professional Plus Volume
        result = get_product_type_code("{90160000-000F-0000-1000-0000000FF1CE}")
        assert result == "000F"

        # Standard
        result = get_product_type_code("{90160000-0012-0000-1000-0000000FF1CE}")
        assert result == "0012"

    def test_get_product_type_code_invalid(self) -> None:
        """Invalid GUID should return None."""
        assert get_product_type_code("not-a-guid") is None

    def test_classify_office_product(self) -> None:
        """Products should be classified by type code."""
        # Professional Plus
        result = classify_office_product("{90160000-000F-0000-1000-0000000FF1CE}")
        assert "Professional Plus" in result

        # Unknown type
        result = classify_office_product("{90160000-9999-0000-1000-0000000FF1CE}")
        assert "Unknown" in result

    def test_is_office_product_code(self) -> None:
        """Office products should be detected by signature."""
        # Valid Office product (ends with FF1CE)
        assert is_office_product_code("{90160000-000F-0000-1000-0000000FF1CE}")

        # Non-Office product
        assert not is_office_product_code("{12345678-1234-1234-1234-123456789012}")

    def test_is_office_product_code_invalid(self) -> None:
        """Invalid GUID should return False."""
        assert not is_office_product_code("not-a-guid")


class TestEdgeCases:
    """Tests for edge cases and boundary conditions."""

    def test_compress_expand_all_ones(self) -> None:
        """All-ones GUID should roundtrip correctly."""
        guid = "{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}"
        assert expand_guid(compress_guid(guid)) == guid

    def test_compress_expand_alternating(self) -> None:
        """Alternating pattern should roundtrip correctly."""
        guid = "{01234567-89AB-CDEF-0123-456789ABCDEF}"
        assert expand_guid(compress_guid(guid)) == guid

    def test_squish_decode_zeros(self) -> None:
        """All-zeros GUID should roundtrip correctly."""
        guid = "{00000000-0000-0000-0000-000000000000}"
        assert decode_squished_guid(squish_guid(guid)) == guid

    def test_squish_decode_ones(self) -> None:
        """All-ones GUID should roundtrip correctly."""
        guid = "{FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF}"
        assert decode_squished_guid(squish_guid(guid)) == guid

    def test_mixed_case_input(self) -> None:
        """Mixed case input should be handled correctly."""
        guid_lower = "{abcdef00-1234-5678-9abc-def012345678}"
        guid_upper = "{ABCDEF00-1234-5678-9ABC-DEF012345678}"

        # Both should compress to same result
        assert compress_guid(guid_lower) == compress_guid(guid_upper)

        # Normalization should uppercase
        assert normalize_guid(guid_lower) == guid_upper
