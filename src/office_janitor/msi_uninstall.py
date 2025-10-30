"""!
@brief Helpers for orchestrating MSI-based Office uninstalls.
@details The routines in this module discover uninstall metadata, compose
`msiexec` command lines, retry failures, and verify registry state to confirm
that the requested product has been removed. Structured telemetry emitted via
:mod:`office_janitor.command_runner` keeps the behaviour aligned with the
reference OffScrub scripts while remaining fully Python-native.
"""
from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Iterable, List, Mapping, MutableMapping, Sequence

from . import command_runner, constants, logging_ext, registry_tools

MSIEXEC_TIMEOUT = 3600
"""!
@brief Maximum seconds to wait for a single ``msiexec`` invocation.
"""

MSIEXEC_BASE_COMMAND = ("msiexec.exe", "/x")
"""!
@brief Command prefix used for MSI product removal.
"""

MSIEXEC_ADDITIONAL_ARGS = ("/qb!", "/norestart")
"""!
@brief UI and reboot suppression arguments mirrored from OffScrub scripts.
"""

MSI_RETRY_ATTEMPTS = 2
"""!
@brief Number of retries after the initial uninstall attempt (total attempts = retries + 1).
"""

MSI_RETRY_DELAY = 5.0
"""!
@brief Seconds to wait between retry attempts.
"""

MSI_VERIFICATION_ATTEMPTS = 3
"""!
@brief Number of registry probes performed when confirming removal.
"""

MSI_VERIFICATION_DELAY = 5.0
"""!
@brief Seconds to wait between registry verification probes.
"""


@dataclass
class _MsiProduct:
    """!
    @brief Normalised metadata describing an MSI product slated for removal.
    """

    product_code: str
    display_name: str
    version: str
    uninstall_handles: Sequence[str]


def _normalise_product_code(raw: str) -> str:
    """!
    @brief Sanitise ``raw`` into the ``{GUID}`` form expected by ``msiexec``.
    """

    token = raw.strip().strip("\0")
    if not token:
        return ""
    core = token.strip("{}")
    if not core:
        return ""
    return f"{{{core.upper()}}}"


def _default_handles_for_code(product_code: str) -> List[str]:
    """!
    @brief Construct registry handle strings for known uninstall roots.
    """

    handles: List[str] = []
    metadata = constants.MSI_PRODUCT_MAP.get(product_code.upper())
    registry_roots = metadata.get("registry_roots", constants.MSI_UNINSTALL_ROOTS) if metadata else constants.MSI_UNINSTALL_ROOTS
    for hive, base in registry_roots:
        handles.append(f"{registry_tools.hive_name(hive)}\\{base}\\{product_code}")
    return handles


def _normalise_product_entry(product: Mapping[str, object] | str) -> _MsiProduct:
    """!
    @brief Convert caller supplied product metadata into an internal structure.
    """

    mapping: MutableMapping[str, object]
    if isinstance(product, MutableMapping):
        mapping = dict(product)
    elif isinstance(product, Mapping):
        mapping = dict(product)
    else:
        mapping = {"product_code": str(product)}

    raw_code = str(
        mapping.get("product_code")
        or mapping.get("ProductCode")
        or mapping.get("code")
        or ""
    )
    product_code = _normalise_product_code(raw_code)
    if not product_code:
        raise ValueError("MSI uninstall entry missing product code")

    properties = mapping.get("properties")
    if isinstance(properties, Mapping):
        property_map = properties
    else:
        property_map = {}

    metadata = constants.MSI_PRODUCT_MAP.get(product_code.upper(), {})
    display_name = str(
        mapping.get("product")
        or property_map.get("display_name")
        or metadata.get("product")
        or product_code
    )
    version = str(
        mapping.get("version")
        or property_map.get("display_version")
        or metadata.get("version")
        or "unknown"
    )

    handles: Sequence[str] = ()
    raw_handles = mapping.get("uninstall_handles")
    if isinstance(raw_handles, Sequence) and not isinstance(raw_handles, (str, bytes)):
        handles = [str(handle).strip() for handle in raw_handles if str(handle).strip()]
    if not handles:
        handles = _default_handles_for_code(product_code)

    return _MsiProduct(
        product_code=product_code,
        display_name=display_name,
        version=version,
        uninstall_handles=tuple(handles),
    )


def _parse_registry_handle(handle: str) -> tuple[int, str] | None:
    """!
    @brief Break a ``HKLM\\...`` style handle into hive/path components.
    """

    cleaned = str(handle).strip()
    if not cleaned or "\\" not in cleaned:
        return None
    prefix, _, path = cleaned.partition("\\")
    hive = constants.REGISTRY_ROOTS.get(prefix.upper())
    if hive is None or not path:
        return None
    return hive, path


def _is_product_present(entry: _MsiProduct) -> bool:
    """!
    @brief Determine whether the product still has uninstall registry entries.
    """

    for handle in entry.uninstall_handles:
        parsed = _parse_registry_handle(handle)
        if parsed and registry_tools.key_exists(parsed[0], parsed[1]):
            return True
    for hive, base in constants.MSI_UNINSTALL_ROOTS:
        if registry_tools.key_exists(hive, f"{base}\\{entry.product_code}"):
            return True
    return False


def build_command(product_code: str) -> List[str]:
    """!
    @brief Compose the ``msiexec`` command used to uninstall ``product_code``.
    """

    normalized = _normalise_product_code(product_code)
    if not normalized:
        raise ValueError("product_code must be a non-empty MSI product code")

    metadata = constants.MSI_PRODUCT_MAP.get(normalized.upper(), {})
    extra_args: Sequence[str] = ()
    potential = metadata.get("msiexec_args") if isinstance(metadata, Mapping) else None
    if isinstance(potential, Sequence) and not isinstance(potential, (str, bytes)):
        extra_args = [str(part) for part in potential if str(part).strip()]
    elif isinstance(potential, str) and potential.strip():
        extra_args = [potential.strip()]

    return [*MSIEXEC_BASE_COMMAND, normalized, *MSIEXEC_ADDITIONAL_ARGS, *extra_args]


def _await_removal(entry: _MsiProduct) -> bool:
    """!
    @brief Poll registry keys to confirm the product has been removed.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    for attempt in range(1, MSI_VERIFICATION_ATTEMPTS + 1):
        present = _is_product_present(entry)
        machine_logger.info(
            "msi_uninstall_verify",
            extra={
                "event": "msi_uninstall_verify",
                "product_code": entry.product_code,
                "attempt": attempt,
                "present": present,
            },
        )
        if not present:
            human_logger.info(
                "Confirmed removal of %s (%s)", entry.display_name, entry.product_code
            )
            return True
        if attempt < MSI_VERIFICATION_ATTEMPTS:
            time.sleep(MSI_VERIFICATION_DELAY)
    return False


def uninstall_products(
    products: Iterable[Mapping[str, object] | str],
    *,
    dry_run: bool = False,
    retries: int = MSI_RETRY_ATTEMPTS,
) -> None:
    """!
    @brief Uninstall the supplied MSI products via ``msiexec``.
    @details Each product is normalised, executed with retry semantics, and
    verified for removal using registry probes. Non-zero exit codes or failed
    verifications raise :class:`RuntimeError` summarising the offending product
    codes.
    @param products Iterable of product codes or inventory mappings.
    @param dry_run When ``True`` log intent without executing ``msiexec``.
    @param retries Additional attempts after the first failure.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    entries: List[_MsiProduct] = []
    for product in products:
        if not product:
            continue
        entries.append(_normalise_product_entry(product))

    if not entries:
        human_logger.info("No MSI products supplied for uninstall; skipping.")
        return

    failures: List[str] = []
    total_attempts = max(1, int(retries) + 1)

    for entry in entries:
        machine_logger.info(
            "msi_uninstall_plan",
            extra={
                "event": "msi_uninstall_plan",
                "product_code": entry.product_code,
                "display_name": entry.display_name,
                "version": entry.version,
                "dry_run": bool(dry_run),
                "handles": list(entry.uninstall_handles),
            },
        )

        if not dry_run and not _is_product_present(entry):
            human_logger.info(
                "%s (%s) is already absent; skipping msiexec.",
                entry.display_name,
                entry.product_code,
            )
            continue

        command = build_command(entry.product_code)
        result: command_runner.CommandResult | None = None

        for attempt in range(1, total_attempts + 1):
            message = (
                "Uninstalling MSI product %s (%s) [attempt %d/%d]"
                % (entry.display_name, entry.product_code, attempt, total_attempts)
            )
            result = command_runner.run_command(
                command,
                event="msi_uninstall",
                timeout=MSIEXEC_TIMEOUT,
                dry_run=dry_run,
                human_message=message,
                extra={
                    "product_code": entry.product_code,
                    "display_name": entry.display_name,
                    "version": entry.version,
                    "attempt": attempt,
                    "attempts": total_attempts,
                },
            )
            if result.skipped:
                break
            if result.returncode == 0:
                break
            if attempt < total_attempts:
                human_logger.warning(
                    "Retrying msiexec for %s (%s)", entry.display_name, entry.product_code
                )
                time.sleep(MSI_RETRY_DELAY)
        else:
            # Loop exhausted without break.
            result = result

        if result is None:
            continue
        if result.skipped or dry_run:
            continue
        if result.returncode != 0:
            machine_logger.error(
                "msi_uninstall_failure",
                extra={
                    "event": "msi_uninstall_failure",
                    "product_code": entry.product_code,
                    "display_name": entry.display_name,
                    "version": entry.version,
                    "return_code": result.returncode,
                },
            )
            failures.append(entry.product_code)
            continue

        if not _await_removal(entry):
            human_logger.error(
                "Registry still reports %s (%s) after msiexec", entry.display_name, entry.product_code
            )
            machine_logger.error(
                "msi_uninstall_residue",
                extra={
                    "event": "msi_uninstall_residue",
                    "product_code": entry.product_code,
                    "display_name": entry.display_name,
                    "version": entry.version,
                },
            )
            failures.append(entry.product_code)

    if failures:
        raise RuntimeError(
            "Failed to uninstall MSI products: %s" % ", ".join(sorted(set(failures)))
        )
