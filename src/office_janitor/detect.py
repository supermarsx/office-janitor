"""!
@brief Detection helpers for installed Microsoft Office components.
@details The detection pipeline queries registry hives, filesystem locations,
and running processes to assemble an inventory of MSI and Click-to-Run Office
deployments as described in the project specification.
"""
from __future__ import annotations

from typing import Dict, List


def detect_msi_installations() -> List[Dict[str, str]]:
    """!
    @brief Inspect the registry and return metadata for MSI-based Office installs.
    """

    raise NotImplementedError


def detect_c2r_installations() -> List[Dict[str, str]]:
    """!
    @brief Probe Click-to-Run configuration to describe installed suites.
    """

    raise NotImplementedError


def gather_office_inventory() -> Dict[str, List[Dict[str, str]]]:
    """!
    @brief Aggregate MSI, C2R, and ancillary signals into an inventory payload.
    """

    raise NotImplementedError
