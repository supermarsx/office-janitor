"""!
@brief Plain console user interface helpers.
@details Provides the interactive menu experience described in the specification
for environments that do not support the richer TUI renderer.
"""
from __future__ import annotations

import textwrap
from typing import Callable, Mapping, MutableMapping


MenuHandler = Callable[[MutableMapping[str, object]], None]


def run_cli(app_state: Mapping[str, object]) -> None:
    """!
    @brief Launch the basic interactive console menu.
    """

    args = app_state.get("args")
    human_logger = app_state.get("human_logger")
    input_func: Callable[[str], str] = app_state.get("input", input)  # type: ignore[assignment]

    if getattr(args, "quiet", False) or getattr(args, "json", False):
        if human_logger:
            human_logger.warning(
                "Interactive menu suppressed because quiet/json output mode was requested."
            )
        return

    detector: Callable[[], Mapping[str, object]] = app_state["detector"]  # type: ignore[assignment]
    planner: Callable[[Mapping[str, object], Mapping[str, object] | None], list[dict]] = app_state["planner"]  # type: ignore[assignment]
    executor: Callable[[list[dict], Mapping[str, object] | None], None] = app_state["executor"]  # type: ignore[assignment]

    menu: list[tuple[str, MenuHandler]] = [
        ("Detect & show installed Office", _menu_detect),
        ("Auto scrub everything detected (recommended)", _menu_auto_all),
        ("Targeted scrub (choose versions/components)", _menu_targeted),
        ("Cleanup only (licenses, residue)", _menu_cleanup),
        ("Diagnostics only (export plan & inventory)", _menu_diagnostics),
        ("Settings (restore point, logging, backups)", _menu_settings),
        ("Exit", _menu_exit),
    ]

    context: MutableMapping[str, object] = {
        "detector": detector,
        "planner": planner,
        "executor": executor,
        "args": args,
        "human_logger": human_logger,
        "input": input_func,
        "inventory": None,
        "plan": None,
        "running": True,
    }

    while context.get("running", True):
        _print_menu(menu)
        selection = input_func("Select an option (1-7): ").strip()
        if not selection.isdigit():
            print("Please enter a number between 1 and 7.")
            continue
        index = int(selection) - 1
        if index < 0 or index >= len(menu):
            print("Please choose a valid menu entry.")
            continue
        handler = menu[index][1]
        try:
            handler(context)
        except Exception as exc:  # pragma: no cover - defensive user feedback
            if human_logger:
                human_logger.error("Menu action failed: %s", exc)
            else:
                print(f"Error: {exc}")


def _print_menu(menu: list[tuple[str, MenuHandler]]) -> None:
    """!
    @brief Render the text menu to stdout.
    """

    header = textwrap.dedent(
        """
        ================= Office Janitor =================
        1. Detect & show installed Office
        2. Auto scrub everything detected (recommended)
        3. Targeted scrub (choose versions/components)
        4. Cleanup only (licenses, residue)
        5. Diagnostics only (export plan & inventory)
        6. Settings (restore point, logging, backups)
        7. Exit
        --------------------------------------------------
        """
    ).strip("\n")
    print(header)


def _menu_detect(context: MutableMapping[str, object]) -> None:
    detector: Callable[[], Mapping[str, object]] = context["detector"]  # type: ignore[assignment]
    inventory = detector()
    context["inventory"] = inventory
    print("Detected inventory:")
    for key, items in inventory.items():
        try:
            count = len(items)  # type: ignore[arg-type]
        except TypeError:
            count = len(list(items))  # type: ignore[arg-type]
        print(f"  - {key}: {count} entries")


def _menu_auto_all(context: MutableMapping[str, object]) -> None:
    _plan_and_execute(context, {"mode": "auto-all"})


def _menu_targeted(context: MutableMapping[str, object]) -> None:
    input_func: Callable[[str], str] = context.get("input", input)  # type: ignore[assignment]
    raw = input_func("Enter comma-separated target versions (e.g. 2016,365): ")
    targets = [item.strip() for item in raw.split(",") if item.strip()]
    overrides: dict[str, object] = {"mode": "target:" + ",".join(targets) if targets else "interactive"}
    if targets:
        overrides["target"] = ",".join(targets)
    _plan_and_execute(context, overrides)


def _menu_cleanup(context: MutableMapping[str, object]) -> None:
    _plan_and_execute(context, {"mode": "cleanup-only", "cleanup_only": True})


def _menu_diagnostics(context: MutableMapping[str, object]) -> None:
    plan_steps = _ensure_plan(context, {"mode": "diagnose", "diagnose": True})
    executor: Callable[[list[dict], Mapping[str, object] | None], None] = context["executor"]  # type: ignore[assignment]
    inventory = context.get("inventory")
    executor(plan_steps, {"mode": "diagnose", "diagnose": True, "inventory": inventory})
    print("Diagnostics captured; no actions executed.")


def _menu_settings(context: MutableMapping[str, object]) -> None:
    args = context.get("args")
    print("Current settings:")
    print(f"  Dry-run: {bool(getattr(args, 'dry_run', False))}")
    print(f"  Create restore point: {not bool(getattr(args, 'no_restore_point', False))}")
    print(f"  Log directory: {getattr(args, 'logdir', '(default)')}")
    print(f"  Backup directory: {getattr(args, 'backup', '(disabled)')}")


def _menu_exit(context: MutableMapping[str, object]) -> None:
    context["running"] = False
    print("Exiting Office Janitor.")


def _plan_and_execute(context: MutableMapping[str, object], overrides: Mapping[str, object]) -> None:
    plan_steps = _ensure_plan(context, overrides)
    executor: Callable[[list[dict], Mapping[str, object] | None], None] = context["executor"]  # type: ignore[assignment]
    payload = dict(overrides)
    if "inventory" not in payload and context.get("inventory") is not None:
        payload["inventory"] = context.get("inventory")
    executor(plan_steps, payload)


def _ensure_plan(context: MutableMapping[str, object], overrides: Mapping[str, object]) -> list[dict]:
    planner: Callable[[Mapping[str, object], Mapping[str, object] | None], list[dict]] = context["planner"]  # type: ignore[assignment]
    inventory = context.get("inventory")
    if inventory is None:
        detector: Callable[[], Mapping[str, object]] = context["detector"]  # type: ignore[assignment]
        inventory = detector()
        context["inventory"] = inventory
    plan_steps = planner(inventory, overrides)
    context["plan"] = plan_steps
    print("Plan contains %d step(s)." % len(plan_steps))
    return plan_steps
