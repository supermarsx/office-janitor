"""!
@brief Native replacements for OffScrub helper flows (full parity)
@details This module provides native Python implementations that mirror the
behaviour and command-line invocation signatures of the legacy OffScrub VBS
helpers. It is intended to be invoked either programmatically via
:func:`uninstall_products` or as a CLI using ``python -m office_janitor.off_scrub_native``.
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path
from typing import Mapping, Sequence, MutableMapping, List

from . import (
    c2r_uninstall,
    msi_uninstall,
    logging_ext,
    registry_tools,
    tasks_services,
    constants,
)


def uninstall_products(config: Mapping[str, object], *, dry_run: bool = False, retries: int | None = None) -> None:
    """!
    @brief Native Click-to-Run uninstall wrapper matching OffScrubC2R behavior.
    @details Reuses the well-tested :mod:`c2r_uninstall` implementation but
    preserves the OffScrub-style logging and options surface so callers/scripts
    can be migrated to this module directly.
    """

    human_logger = logging_ext.get_human_logger()
    human_logger.info("OffScrub native C2R: starting uninstall (dry_run=%s)", bool(dry_run))

    kwargs = {"dry_run": dry_run}
    if retries is not None:
        kwargs["retries"] = retries

    # Delegate to the existing implementation which already performs the
    # required service stops, invocation of OfficeC2RClient or setup.exe, and
    # verification probing.
    c2r_uninstall.uninstall_products(config, **kwargs)


def uninstall_msi_products(products: Sequence[Mapping[str, object] | str], *, dry_run: bool = False, retries: int | None = None) -> None:
    """!
    @brief Native MSI OffScrub entry point.
    @details Mirrors the semantics of the VBS MSI helpers by calling into the
    existing :mod:`msi_uninstall` module and preserving logging and retry
    semantics.
    """

    human_logger = logging_ext.get_human_logger()
    human_logger.info("OffScrub native MSI: uninstalling %d products (dry_run=%s)", len(list(products)), bool(dry_run))

    kwargs = {"dry_run": dry_run}
    if retries is not None:
        kwargs["retries"] = retries

    msi_uninstall.uninstall_products(products, **kwargs)


def _parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(prog="off_scrub_native")
    sub = parser.add_subparsers(dest="command")

    c2r = sub.add_parser("c2r", help="Click-to-Run uninstall wrapper")
    c2r.add_argument("--dry-run", action="store_true", dest="dry_run")
    c2r.add_argument("--retries", type=int, default=None)
    c2r.add_argument("--release-ids", nargs="*", default=None)
    c2r.add_argument("--display-name", default=None)

    msi = sub.add_parser("msi", help="MSI uninstall wrapper")
    msi.add_argument("--dry-run", action="store_true", dest="dry_run")
    msi.add_argument("--retries", type=int, default=None)
    msi.add_argument("--product-codes", nargs="*", default=None)

    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    """CLI entrypoint to mimic OffScrub script behaviour.

    Usage examples:
    - python -m office_janitor.off_scrub_native c2r --release-ids PRODUCTION
    - python -m office_janitor.off_scrub_native msi --product-codes {GUID}
    """

    args = _parse_args(argv)
    human_logger = logging_ext.get_human_logger()

    try:
        if args.command == "c2r":
            config: MutableMapping[str, object] = {}
            if args.release_ids:
                config["release_ids"] = args.release_ids
            if args.display_name:
                config["product"] = args.display_name
            uninstall_products(config, dry_run=bool(args.dry_run), retries=args.retries)
            return 0
        elif args.command == "msi":
            products: List[Mapping[str, object] | str] = []
            if args.product_codes:
                for code in args.product_codes:
                    products.append({"product_code": code})
            uninstall_msi_products(products, dry_run=bool(args.dry_run), retries=args.retries)
            return 0
        else:
            human_logger.info("No command supplied; nothing to do.")
            return 2
    except Exception as exc:  # pragma: no cover - propagate to caller
        human_logger.error("OffScrub native operation failed: %s", exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
