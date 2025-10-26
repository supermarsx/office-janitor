"""!
@brief Helpers for orchestrating MSI-based Office uninstalls.
@details Translates the OffScrub command sequences present in
``OfficeScrubber.cmd`` into Python helpers that invoke the matching VBS
automation scripts with the same argument conventions. This allows the
scrubber to mimic the legacy workflow while still benefiting from structured
logging and dry-run enforcement.
"""
from __future__ import annotations

import subprocess
import time
from pathlib import Path
from typing import Iterable, List, Mapping, MutableMapping

from . import logging_ext
from .off_scrub_scripts import ensure_offscrub_script

CSCRIPT = "cscript.exe"
"""!
@brief Host executable used to run OffScrub VBS helpers.
"""

OFFSCRUB_TIMEOUT = 1800
"""!
@brief Default timeout (seconds) for OffScrub automation helpers.
"""

OFFSCRUB_BASE_ARGS: tuple[str, ...] = (
    "ALL",
    "/OSE",
    "/NOCANCEL",
    "/FORCE",
    "/ENDCURRENTINSTALLS",
    "/DELETEUSERSETTINGS",
    "/CLEARADDINREG",
    "/REMOVELYNC",
)
"""!
@brief Argument list mirrored from ``_para`` in ``OfficeScrubber.cmd``.
"""

OFFSCRUB_SCRIPT_MAP: Mapping[str, str] = {
    "2003": "OffScrub03.vbs",
    "2007": "OffScrub07.vbs",
    "2010": "OffScrub10.vbs",
    "2013": "OffScrub_O15msi.vbs",
    "2016": "OffScrub_O16msi.vbs",
    "2019": "OffScrub_O16msi.vbs",
    "2021": "OffScrub_O16msi.vbs",
    "2024": "OffScrub_O16msi.vbs",
    "365": "OffScrub_O16msi.vbs",
}
"""!
@brief Mapping between detected Office versions and OffScrub MSI helpers.
"""

DEFAULT_OFFSCRUB_SCRIPT = "OffScrub_O16msi.vbs"
"""!
@brief Fallback helper used when a specific version is not known.
"""


def _sanitize_product_code(product_code: str) -> str:
    """!
    @brief Produce a filesystem-safe representation of ``product_code``.
    """

    return product_code.strip().strip("{}").replace("-", "").upper() or "unknown"


def _resolve_offscrub_script(version_hint: str | None) -> str:
    """!
    @brief Select the correct OffScrub helper given an optional version hint.
    """

    if version_hint:
        normalized = version_hint.strip().lower()
        mapped = OFFSCRUB_SCRIPT_MAP.get(normalized)
        if mapped:
            return mapped
        for key, value in OFFSCRUB_SCRIPT_MAP.items():
            if normalized.startswith(key.lower()):
                return value
    return DEFAULT_OFFSCRUB_SCRIPT


def build_command(
    product: Mapping[str, object] | str,
    *,
    script_directory: Path | None = None,
) -> List[str]:
    """!
    @brief Compose the OffScrub command line for a given MSI installation.
    @details The helper accepts either a plain product code or the inventory
    record emitted by :mod:`detect`. Only fields understood by the OffScrub
    family are translated into command switches so that the runtime matches the
    reference batch script.
    """

    if isinstance(product, MutableMapping):
        product_code = str(product.get("product_code", "")).strip()
        version_hint = str(product.get("version", "")).strip()
    elif isinstance(product, Mapping):
        product_code = str(product.get("product_code", "")).strip()
        version_hint = str(product.get("version", "")).strip()
    else:
        product_code = str(product).strip()
        version_hint = ""

    script_name = _resolve_offscrub_script(version_hint)
    script_path = ensure_offscrub_script(script_name, base_directory=script_directory)

    command: List[str] = [str(CSCRIPT), "//NoLogo", str(script_path)]
    command.extend(OFFSCRUB_BASE_ARGS)
    if product_code:
        command.append(f"/PRODUCTCODE={product_code}")
    return command


def uninstall_products(
    products: Iterable[Mapping[str, object] | str],
    *,
    dry_run: bool = False,
) -> None:
    """!
    @brief Uninstall the supplied MSI products using OffScrub semantics.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    entries: List[Mapping[str, object] | str] = [product for product in products if product]
    if not entries:
        human_logger.info("No MSI products supplied for uninstall; skipping.")
        return

    human_logger.info(
        "Preparing to run OffScrub for %d MSI product(s). Dry-run: %s",
        len(entries),
        bool(dry_run),
    )

    failures: List[str] = []

    for product in entries:
        command = build_command(product)
        if isinstance(product, Mapping):
            product_code = str(product.get("product_code", "")).strip()
            version_hint = str(product.get("version", "")).strip()
        else:
            product_code = str(product).strip()
            version_hint = ""

        safe_code = _sanitize_product_code(product_code or version_hint or "")
        machine_logger.info(
            "msi_uninstall_plan",
            extra={
                "event": "msi_uninstall_plan",
                "product_code": product_code or None,
                "version_hint": version_hint or None,
                "dry_run": bool(dry_run),
                "command": command,
            },
        )

        if dry_run:
            human_logger.info("Dry-run: would invoke %s", " ".join(command))
            continue

        human_logger.info(
            "Invoking OffScrub helper %s for MSI target %s",
            command[2],
            product_code or version_hint or safe_code,
        )
        start = time.monotonic()
        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=OFFSCRUB_TIMEOUT,
                check=False,
            )
        except subprocess.TimeoutExpired as exc:
            duration = time.monotonic() - start
            human_logger.error(
                "OffScrub timed out after %.1fs for %s", duration, product_code or safe_code
            )
            machine_logger.error(
                "msi_uninstall_timeout",
                extra={
                    "event": "msi_uninstall_timeout",
                    "product_code": product_code or None,
                    "version_hint": version_hint or None,
                    "duration": duration,
                    "stdout": exc.stdout,
                    "stderr": exc.stderr,
                },
            )
            failures.append(product_code or safe_code)
            continue

        duration = time.monotonic() - start
        if result.returncode != 0:
            human_logger.error(
                "OffScrub helper for %s failed with exit code %s",
                product_code or safe_code,
                result.returncode,
            )
            machine_logger.error(
                "msi_uninstall_failure",
                extra={
                    "event": "msi_uninstall_failure",
                    "product_code": product_code or None,
                    "version_hint": version_hint or None,
                    "return_code": result.returncode,
                    "stdout": result.stdout,
                    "stderr": result.stderr,
                    "duration": duration,
                },
            )
            failures.append(product_code or safe_code)
            continue

        human_logger.info(
            "Successfully completed OffScrub for %s in %.1f seconds.",
            product_code or safe_code,
            duration,
        )
        machine_logger.info(
            "msi_uninstall_success",
            extra={
                "event": "msi_uninstall_success",
                "product_code": product_code or None,
                "version_hint": version_hint or None,
                "return_code": result.returncode,
                "stdout": result.stdout,
                "stderr": result.stderr,
                "duration": duration,
            },
        )

    if failures:
        raise RuntimeError(f"Failed to uninstall MSI products: {', '.join(failures)}")
