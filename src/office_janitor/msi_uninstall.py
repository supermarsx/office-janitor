"""!
@brief Helpers for orchestrating MSI-based Office uninstalls.
@details The routines in this module discover uninstall metadata, compose
`msiexec` command lines, retry failures, and verify registry state to confirm
that the requested product has been removed. Structured telemetry emitted via
:mod:`office_janitor.command_runner` keeps the behaviour aligned with the
reference OffScrub scripts while remaining fully Python-native.
"""

from __future__ import annotations

import os
import shlex
import sys
import time
from collections.abc import Iterable, Mapping, MutableMapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

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

MSI_BUSY_RETURN_CODE = 1618
"""!
@brief ``msiexec`` exit code indicating another installation is already running.
"""

MSI_BUSY_BACKOFF_CAP = 60.0
"""!
@brief Maximum seconds to wait before retrying when Windows Installer is busy.
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
    maintenance_executable: str | None = None


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


def _default_handles_for_code(product_code: str) -> list[str]:
    """!
    @brief Construct registry handle strings for known uninstall roots.
    """

    handles: list[str] = []
    metadata = constants.MSI_PRODUCT_MAP.get(product_code.upper())
    registry_roots = (
        metadata.get("registry_roots", constants.MSI_UNINSTALL_ROOTS)
        if metadata
        else constants.MSI_UNINSTALL_ROOTS
    )
    for hive, base in registry_roots:
        handles.append(f"{registry_tools.hive_name(hive)}\\{base}\\{product_code}")
    return handles


def _strip_icon_index(raw: str) -> str:
    """!
    @brief Remove trailing icon index fragments from registry values.
    """

    text = raw.strip()
    if "," not in text:
        return text
    prefix, _, suffix = text.partition(",")
    remainder = suffix.strip()
    if remainder and not remainder.lstrip("+-").isdigit():
        return text
    return prefix


def _extract_setup_candidate(value: object) -> str:
    """!
    @brief Attempt to extract a ``setup.exe`` path from ``value``.
    """

    if value is None:
        return ""
    text = _strip_icon_index(str(value).strip())
    if not text or "setup.exe" not in text.lower():
        return ""
    cleaned_text = text.strip().strip('"').strip()
    candidates: list[str] = []
    if cleaned_text:
        candidates.append(cleaned_text)
    try:
        parts = shlex.split(text, posix=False)
    except ValueError:
        parts = []
    if parts:
        candidates.extend(parts)
    for token in candidates:
        cleaned = token.strip().strip('"').strip()
        if not cleaned:
            continue
        lowered = cleaned.lower()
        if lowered.endswith("setup.exe"):
            return cleaned
        if "setup.exe" in lowered:
            index = lowered.find("setup.exe")
            return cleaned[: index + len("setup.exe")]
    return ""


def _normalise_maintenance_paths(*sources: object) -> tuple[str, ...]:
    """!
    @brief Normalise a collection of maintenance path hints into unique strings.
    """

    collected: list[str] = []
    for source in sources:
        if not source:
            continue
        if isinstance(source, (list, tuple, set)):
            for item in source:
                candidate = _extract_setup_candidate(item)
                if candidate and candidate not in collected:
                    collected.append(candidate)
        else:
            candidate = _extract_setup_candidate(source)
            if candidate and candidate not in collected:
                collected.append(candidate)
    return tuple(collected)


def _select_existing_setup(paths: Sequence[str]) -> str | None:
    """!
    @brief Return the first existing ``setup.exe`` from ``paths``.
    """

    for raw in paths:
        candidate = str(raw).strip().strip('"')
        if not candidate:
            continue
        expanded = os.path.expanduser(os.path.expandvars(candidate))
        path = Path(expanded)
        try:
            if path.is_file() and path.name.lower() == "setup.exe":
                return str(path)
        except OSError:
            continue
    return None


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
        mapping.get("product_code") or mapping.get("ProductCode") or mapping.get("code") or ""
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

    maintenance_paths = _normalise_maintenance_paths(
        mapping.get("maintenance_paths"),
        property_map.get("maintenance_paths"),
        mapping.get("display_icon"),
        property_map.get("display_icon"),
        property_map.get("uninstall_string"),
    )
    maintenance_executable = _select_existing_setup(maintenance_paths)

    return _MsiProduct(
        product_code=product_code,
        display_name=display_name,
        version=version,
        uninstall_handles=tuple(handles),
        maintenance_executable=maintenance_executable,
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


def _compute_busy_backoff(attempt: int) -> float:
    """!
    @brief Calculate a retry delay for busy Windows Installer states.
    @details Implements exponential backoff capped at :data:`MSI_BUSY_BACKOFF_CAP`
    so repeated ``ERROR_INSTALL_ALREADY_RUNNING`` responses do not hammer the
    service while still progressing promptly once the conflicting installer
    exits.
    @param attempt Current attempt number (1-indexed).
    @returns Delay in seconds before the next retry.
    """

    exponent = max(0, int(attempt) - 1)
    delay = MSI_RETRY_DELAY * (2**exponent)
    return float(min(MSI_BUSY_BACKOFF_CAP, delay))


def _handle_busy_installer(
    entry: _MsiProduct,
    *,
    attempt: int,
    attempts: int,
    input_func: Callable[[str], str] | None = None,
) -> tuple[bool, float]:
    """!
    @brief Emit guidance and optionally prompt when Windows Installer is busy.
    @details ``msiexec`` returns ``1618`` when another installation is already
    running. This helper surfaces actionable guidance through both human and
    structured logs, then decides whether to retry the uninstall based on
    operator confirmation or automation context. When retrying it returns the
    computed backoff delay so callers can pause before the next attempt.
    @param entry Normalised product metadata for the pending uninstall.
    @param attempt Current attempt number (1-indexed).
    @param attempts Total attempts allowed for the uninstall.
    @param input_func Optional override for collecting operator responses in
    interactive sessions.
    @returns Tuple of ``(should_retry, delay_seconds)``.
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    delay = _compute_busy_backoff(attempt)
    interactive = False

    if input_func is not None:
        interactive = True
    else:
        stdin = getattr(sys, "stdin", None)
        isatty = getattr(stdin, "isatty", None)
        interactive = bool(isatty and isatty())
        if interactive:
            input_func = input

    human_logger.warning(
        (
            "Windows Installer reported another setup already running while "
            "removing %s (%s). Close any other installers or finish updates "
            "before retrying."
        ),
        entry.display_name,
        entry.product_code,
    )

    decision = "retry"

    if not interactive:
        human_logger.info(
            "Non-interactive session detected; automatically retrying after %.1fs.",
            delay,
        )
    else:
        prompt = (
            f"Retry uninstall of {entry.display_name} after waiting {delay:.0f}s once other "
            "installers are closed? [Y/n]: "
        )
        try:
            response = input_func(prompt)
        except EOFError:
            response = "n"
        normalized = response.strip().lower()
        if normalized in {"n", "no"}:
            decision = "cancel"
            human_logger.info(
                "Operator cancelled uninstall retries for %s (%s) due to busy installer.",
                entry.display_name,
                entry.product_code,
            )
            machine_logger.info(
                "msi_uninstall_busy",
                extra={
                    "event": "msi_uninstall_busy",
                    "product_code": entry.product_code,
                    "display_name": entry.display_name,
                    "version": entry.version,
                    "attempt": attempt,
                    "attempts": attempts,
                    "decision": decision,
                    "delay": delay,
                    "interactive": interactive,
                },
            )
            return False, 0.0
        human_logger.info(
            "Operator approved retry once other installers exit (delay %.1fs).",
            delay,
        )

    machine_logger.info(
        "msi_uninstall_busy",
        extra={
            "event": "msi_uninstall_busy",
            "product_code": entry.product_code,
            "display_name": entry.display_name,
            "version": entry.version,
            "attempt": attempt,
            "attempts": attempts,
            "decision": decision,
            "delay": delay,
            "interactive": interactive,
        },
    )
    return True, delay


def build_command(product_code: str, *, maintenance_executable: str | None = None) -> list[str]:
    """!
    @brief Compose the command used to uninstall ``product_code``.
    """

    if maintenance_executable:
        cleaned = maintenance_executable.strip()
        if not cleaned:
            raise ValueError("maintenance_executable must be non-empty when provided")
        # For setup.exe, use /uninstall with the product ID
        # VBS equivalent: setup.exe /uninstall <ProductID> /dll OSETUP.DLL
        return [cleaned, "/uninstall", product_code]

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


# ---------------------------------------------------------------------------
# Setup.exe Uninstall Orchestration (VBS parity)
# ---------------------------------------------------------------------------

SETUP_UNINSTALL_CONFIG_TEMPLATE = """<Configuration Product="{product_id}">
  <Display Level="none" CompletionNotice="no" SuppressModal="yes" AcceptEula="yes" />
  <Logging Type="verbose" Path="%temp%" />
  <Setting Id="SETUP_REBOOT" Value="Never" />
  <Setting Id="SETUP_REBOOT_PROMPT" Value="Never" />
</Configuration>
"""
"""!
@brief Template for generating setup.exe configuration XML.
@details VBS equivalent: config.xml generation in SetupExeRemoval subroutine.
"""


def build_setup_config_xml(
    product_id: str,
    output_path: Path | str | None = None,
) -> Path:
    """!
    @brief Generate a configuration XML file for setup.exe uninstall.
    @details VBS equivalent: config.xml generation in O15/O16 MSI scripts.
    @param product_id The Office product ID (e.g., "ProPlus", "Standard", etc.)
    @param output_path Where to write the XML. Defaults to temp directory.
    @returns Path to the generated XML file.
    """
    import tempfile

    content = SETUP_UNINSTALL_CONFIG_TEMPLATE.format(product_id=product_id)

    if output_path is None:
        output_path = Path(tempfile.gettempdir()) / "OJUninstallConfig.xml"
    else:
        output_path = Path(output_path)

    output_path.write_text(content, encoding="utf-8")
    return output_path


def find_setup_exe_from_registry(product_code: str) -> Path | None:
    """!
    @brief Locate setup.exe from registry InstallSource or InstallLocation.
    @details VBS equivalent: setup.exe discovery in SetupExeRemoval.
    @param product_code The MSI product code GUID.
    @returns Path to setup.exe if found, None otherwise.
    """
    human_logger = logging_ext.get_human_logger()
    normalized = _normalise_product_code(product_code)

    # Check Uninstall registry keys
    for hive, root in constants.MSI_UNINSTALL_ROOTS:
        key_path = f"{root}\\{normalized}"
        values = registry_tools.read_values(hive, key_path)
        if not values:
            continue

        # Try InstallSource first
        install_source = values.get("InstallSource", "")
        if install_source and isinstance(install_source, str):
            setup_path = Path(install_source) / "setup.exe"
            if setup_path.exists():
                human_logger.debug("Found setup.exe via InstallSource: %s", setup_path)
                return setup_path

        # Try InstallLocation
        install_location = values.get("InstallLocation", "")
        if install_location and isinstance(install_location, str):
            setup_path = Path(install_location) / "setup.exe"
            if setup_path.exists():
                human_logger.debug("Found setup.exe via InstallLocation: %s", setup_path)
                return setup_path

    return None


def attempt_setup_exe_removal(
    product_code: str,
    product_id: str,
    *,
    dry_run: bool = False,
    timeout: int = MSIEXEC_TIMEOUT,
) -> bool:
    """!
    @brief Try Office setup.exe for cleaner uninstall before msiexec fallback.
    @details VBS equivalent: SetupExeRemoval subroutine in O15/O16 scripts.
    Steps:
    1. Locate setup.exe from InstallSource or InstallLocation registry
    2. Build uninstall config XML
    3. Execute: setup.exe /uninstall <ProductID> /config <xml>
    4. Return True if successful, False to try msiexec
    @param product_code The MSI product code GUID.
    @param product_id The Office product ID (e.g., "ProPlus").
    @param dry_run If True, only log what would be done.
    @param timeout Command timeout in seconds.
    @returns True if uninstall succeeded, False otherwise.
    """
    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    # Try to find setup.exe
    setup_exe = find_setup_exe_from_registry(product_code)
    if not setup_exe:
        human_logger.debug("setup.exe not found for product %s", product_code)
        return False

    # Build config XML
    config_xml = build_setup_config_xml(product_id)

    command = [
        str(setup_exe),
        "/uninstall",
        product_id,
        "/config",
        str(config_xml),
    ]

    if dry_run:
        human_logger.info("[DRY-RUN] Would execute: %s", " ".join(command))
        return True

    human_logger.info("Attempting setup.exe removal for %s", product_id)
    machine_logger.info(
        "msi_setup_exe_removal_start",
        extra={
            "event": "msi_setup_exe_removal_start",
            "product_code": product_code,
            "product_id": product_id,
            "setup_exe": str(setup_exe),
            "config_xml": str(config_xml),
        },
    )

    result = command_runner.run_command(
        command,
        event="msi_setup_exe_uninstall",
        timeout=timeout,
    )

    machine_logger.info(
        "msi_setup_exe_removal_complete",
        extra={
            "event": "msi_setup_exe_removal_complete",
            "product_code": product_code,
            "product_id": product_id,
            "exit_code": result.returncode,
        },
    )

    if result.returncode == 0:
        human_logger.info("setup.exe removal completed successfully for %s", product_id)
        return True

    human_logger.warning("setup.exe removal returned %d for %s", result.returncode, product_id)
    return False


def _run_uninstall_command(
    entry: _MsiProduct,
    command: Sequence[str],
    *,
    using_setup: bool,
    total_attempts: int,
    dry_run: bool,
    busy_input_func: Callable[[str], str] | None,
) -> command_runner.CommandResult | None:
    """!
    @brief Execute ``command`` with retry semantics for ``entry``.
    @details Mirrors the legacy OffScrub retry loop while allowing callers to
    swap between ``msiexec`` and ``setup.exe`` executors. ``msiexec`` commands
    continue to receive busy-installer handling whereas setup-based fallbacks do
    not since they are external bootstrappers.
    @returns The final :class:`~office_janitor.command_runner.CommandResult` or
    ``None`` if execution was skipped entirely.
    """

    human_logger = logging_ext.get_human_logger()

    event_name = "msi_setup_uninstall" if using_setup else "msi_uninstall"
    result: command_runner.CommandResult | None = None

    for attempt in range(1, total_attempts + 1):
        if using_setup:
            message = (
                f"Uninstalling MSI product {entry.display_name} ({entry.product_code}) "
                f"via setup.exe [attempt {attempt}/{total_attempts}]"
            )
        else:
            message = (
                f"Uninstalling MSI product {entry.display_name} ({entry.product_code}) "
                f"[attempt {attempt}/{total_attempts}]"
            )

        result = command_runner.run_command(
            list(command),
            event=event_name,
            timeout=MSIEXEC_TIMEOUT,
            dry_run=dry_run,
            human_message=message,
            extra={
                "product_code": entry.product_code,
                "display_name": entry.display_name,
                "version": entry.version,
                "attempt": attempt,
                "attempts": total_attempts,
                "executor": command[0] if command else None,
            },
        )

        if result.skipped:
            break
        if result.returncode == 0:
            break
        if dry_run:
            break

        if (
            not using_setup
            and result.returncode == MSI_BUSY_RETURN_CODE
            and attempt < total_attempts
        ):
            should_retry, delay = _handle_busy_installer(
                entry,
                attempt=attempt,
                attempts=total_attempts,
                input_func=busy_input_func,
            )
            if should_retry:
                if delay > 0:
                    time.sleep(delay)
                continue
            break

        if attempt < total_attempts:
            human_logger.warning(
                "Retrying %s for %s (%s)",
                command[0] if command else "uninstall",
                entry.display_name,
                entry.product_code,
            )
            time.sleep(MSI_RETRY_DELAY)

    return result


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
    busy_input_func: Callable[[str], str] | None = None,
) -> None:
    """!
    @brief Uninstall the supplied MSI products via ``msiexec`` or setup fallbacks.
    @details Each product is normalised, executed with retry semantics, and
    verified for removal using registry probes. Non-zero exit codes or failed
    verifications raise :class:`RuntimeError` summarising the offending product
    codes.
    @param products Iterable of product codes or inventory mappings.
    @param dry_run When ``True`` log intent without executing ``msiexec``.
    @param retries Additional attempts after the first failure.
    @param busy_input_func Optional callback used when prompting about busy
    Windows Installer sessions (exit code ``1618``).
    """

    human_logger = logging_ext.get_human_logger()
    machine_logger = logging_ext.get_machine_logger()

    entries: list[_MsiProduct] = []
    for product in products:
        if not product:
            continue
        entries.append(_normalise_product_entry(product))

    if not entries:
        human_logger.info("No MSI products supplied for uninstall; skipping.")
        return

    failures: list[str] = []
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
                "maintenance_executable": entry.maintenance_executable,
            },
        )

        if not dry_run and not _is_product_present(entry):
            human_logger.info(
                "%s (%s) is already absent; skipping msiexec.",
                entry.display_name,
                entry.product_code,
            )
            continue

        primary_command = build_command(entry.product_code)
        result = _run_uninstall_command(
            entry,
            primary_command,
            using_setup=False,
            total_attempts=total_attempts,
            dry_run=dry_run,
            busy_input_func=busy_input_func,
        )
        command: Sequence[str] = primary_command

        if (
            not dry_run
            and result is not None
            and not result.skipped
            and result.returncode != 0
            and entry.maintenance_executable
        ):
            fallback_command = build_command(
                entry.product_code, maintenance_executable=entry.maintenance_executable
            )
            human_logger.warning(
                ("msiexec uninstall of %s (%s) returned %d; attempting setup.exe fallback"),
                entry.display_name,
                entry.product_code,
                result.returncode,
            )
            machine_logger.info(
                "msi_uninstall_fallback",
                extra={
                    "event": "msi_uninstall_fallback",
                    "product_code": entry.product_code,
                    "display_name": entry.display_name,
                    "version": entry.version,
                    "return_code": result.returncode,
                    "fallback_executable": entry.maintenance_executable,
                },
            )
            result = _run_uninstall_command(
                entry,
                fallback_command,
                using_setup=True,
                total_attempts=total_attempts,
                dry_run=dry_run,
                busy_input_func=busy_input_func,
            )
            command = fallback_command

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
                    "executor": command[0] if command else None,
                },
            )
            failures.append(entry.product_code)
            continue

        if not _await_removal(entry):
            human_logger.error(
                "Registry still reports %s (%s) after uninstall command",
                entry.display_name,
                entry.product_code,
            )
            machine_logger.error(
                "msi_uninstall_residue",
                extra={
                    "event": "msi_uninstall_residue",
                    "product_code": entry.product_code,
                    "display_name": entry.display_name,
                    "version": entry.version,
                    "executor": command[0] if command else None,
                },
            )
            failures.append(entry.product_code)

    if failures:
        raise RuntimeError(
            "Failed to uninstall MSI products: {}".format(", ".join(sorted(set(failures))))
        )
