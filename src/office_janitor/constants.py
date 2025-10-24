"""!
@brief Static data and enumerations for Office Janitor.
@details Holds product code mappings, registry roots, default paths, and other
constants used across detection, planning, and scrub orchestration per the
specification.
"""
from __future__ import annotations

SUPPORTED_VERSIONS = (
    "2003",
    "2007",
    "2010",
    "2013",
    "2016",
    "2019",
    "2021",
    "2024",
    "365",
)

DEFAULT_OFFICE_PROCESSES = (
    "winword.exe",
    "excel.exe",
    "outlook.exe",
    "onenote.exe",
    "visio.exe",
    "powerpnt.exe",
)
