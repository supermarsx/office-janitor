import sys
from office_janitor import c2r_uninstall, licensing


def test_c2r_derives_uninstall_handles():
    config = {"release_ids": ["O365ProPlusRetail"]}
    target = c2r_uninstall._normalise_c2r_entry(config)
    assert target.uninstall_handles, "Expected derived uninstall handles"
    # Support both canonical HKLM/HKCU strings and numeric hive fallbacks (hex)
    assert any(h.startswith("HKLM\\") or h.startswith("HKCU\\") or h.startswith("0x") for h in target.uninstall_handles)
    # Ensure at least one handle references ClickToRun/ProductReleaseIDs/Office
    assert any("ProductReleaseIDs" in h or "ClickToRun" in h or "Office" in h for h in target.uninstall_handles)


def test_parse_license_results():
    out = "OSPP:2\nSPP:3\nSome unrelated line\n"
    counts = licensing._parse_license_results(out)
    assert counts["ospp"] == 2
    assert counts["spp"] == 3
