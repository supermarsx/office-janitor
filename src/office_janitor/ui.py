"""!
@brief Plain console user interface helpers.
@details Implements the numbered text menu described in the specification and
bridges menu actions to the detector/planner/executor callables exposed via the
``app_state`` mapping assembled in :mod:`office_janitor.main`.
"""

from __future__ import annotations

import logging
import textwrap
from collections.abc import Iterable, Mapping, MutableMapping
from typing import Callable, cast

from . import confirm, version
from . import plan as plan_module
from .app_state import AppState

_DEFAULT_MENU_LABELS = [
    "Detect & show installed Office",
    "Auto scrub everything detected (recommended)",
    "Targeted scrub (choose versions/components)",
    "Cleanup only (licenses, residue)",
    "Diagnostics only (export plan & inventory)",
    "Settings (restore point, logging, backups)",
    "Exit",
]


MenuHandler = Callable[[MutableMapping[str, object]], None]


def run_cli(app_state: AppState) -> None:
    """!
    @brief Launch the basic interactive console menu.
    @details The function mirrors the layout in the specification, wiring menu
    selections to the reusable detection/planning/scrubbing pipelines while also
    reflecting runtime settings that can be toggled inside the menu.
    """

    args = app_state["args"]
    human_logger = app_state["human_logger"]
    machine_logger = app_state["machine_logger"]
    emit_event = app_state["emit_event"]
    event_queue = app_state["event_queue"]
    input_func = cast(Callable[[str], str], app_state.get("input", input))

    if getattr(args, "quiet", False) or getattr(args, "json", False):
        if human_logger:
            human_logger.warning(
                "Interactive menu suppressed because quiet/json output mode was requested.",
            )
        if callable(emit_event):
            emit_event(
                "ui.suppressed",
                message="Interactive menu suppressed by CLI flags.",
                mode="cli",
            )
        return

    detector = app_state["detector"]
    planner = app_state["planner"]
    executor = app_state["executor"]

    menu: list[tuple[str, MenuHandler]] = [
        (_DEFAULT_MENU_LABELS[0], _menu_detect),
        (_DEFAULT_MENU_LABELS[1], _menu_auto_all),
        (_DEFAULT_MENU_LABELS[2], _menu_targeted),
        (_DEFAULT_MENU_LABELS[3], _menu_cleanup),
        (_DEFAULT_MENU_LABELS[4], _menu_diagnostics),
        (_DEFAULT_MENU_LABELS[5], _menu_settings),
        (_DEFAULT_MENU_LABELS[6], _menu_exit),
    ]

    context: MutableMapping[str, object] = {
        "detector": detector,
        "planner": planner,
        "executor": executor,
        "args": args,
        "human_logger": human_logger,
        "machine_logger": machine_logger,
        "emit_event": emit_event,
        "event_queue": event_queue,
        "input": input_func,
        "confirm": app_state.get("confirm", confirm.request_scrub_confirmation),
        "inventory": None,
        "plan": None,
        "running": True,
    }

    _notify(context, "ui.start", "Interactive CLI started.")

    while context.get("running", True):
        _print_menu(menu)
        selection = input_func("Select an option (1-7): ").strip()
        if not selection.isdigit():
            _notify(
                context,
                "ui.invalid",
                f"Menu selection {selection!r} is not a number.",
                level="warning",
            )
            print("Please enter a number between 1 and 7.")
            continue
        index = int(selection) - 1
        if index < 0 or index >= len(menu):
            _notify(
                context,
                "ui.invalid",
                f"Menu selection {selection!r} outside valid range.",
                level="warning",
            )
            print("Please choose a valid menu entry.")
            continue
        label, handler = menu[index]
        _notify(context, "ui.select", f"Selected menu option: {label}", index=index + 1)
        try:
            handler(context)
        except Exception as exc:  # pragma: no cover - defensive user feedback
            _notify(context, "ui.error", f"Menu action failed: {exc}", level="error")
            if human_logger:
                human_logger.exception("Menu action failure", exc_info=exc)
            else:
                print(f"Error: {exc}")


def _print_menu(menu: list[tuple[str, MenuHandler]]) -> None:
    """!
    @brief Render the text menu to stdout.
    """

    metadata = version.build_info()
    labels = [entry[0] for entry in menu]
    if not labels:
        labels = list(_DEFAULT_MENU_LABELS)
    elif len(labels) < len(_DEFAULT_MENU_LABELS):
        labels = labels + _DEFAULT_MENU_LABELS[len(labels) :]
    else:
        labels = labels[: len(_DEFAULT_MENU_LABELS)]

    header = textwrap.dedent(
        f"""
        ================= Office Janitor =================
        Version {metadata['version']} (build {metadata['build']})
        --------------------------------------------------
        1. {labels[0]}
        2. {labels[1]}
        3. {labels[2]}
        4. {labels[3]}
        5. {labels[4]}
        6. {labels[5]}
        7. {labels[6]}
        --------------------------------------------------
        """
    ).strip("\n")
    print(header)


def _menu_detect(context: MutableMapping[str, object]) -> None:
    detector = cast(Callable[[], Mapping[str, object]], context["detector"])
    _notify(context, "detect.start", "Starting detection run from CLI menu.")
    inventory = detector()
    context["inventory"] = inventory
    summary = _summarize_inventory(inventory)
    _notify(context, "detect.complete", "Detection run finished.", inventory=summary)
    print("Detected inventory:")
    for key, count in summary.items():
        print(f"  - {key}: {count} entries")


def _menu_auto_all(context: MutableMapping[str, object]) -> None:
    _plan_and_execute(context, {"mode": "auto-all", "auto_all": True}, label="auto scrub")


def _menu_targeted(context: MutableMapping[str, object]) -> None:
    input_func = cast(Callable[[str], str], context.get("input", input))
    raw = input_func("Enter comma-separated target versions (e.g. 2016,365): ")
    targets = [item.strip() for item in raw.split(",") if item.strip()]
    includes_raw = input_func(
        "Optional: include additional components (visio,project,onenote): "
    ).strip()
    overrides: dict[str, object] = {
        "mode": "target:" + ",".join(targets) if targets else "interactive"
    }
    if targets:
        joined = ",".join(targets)
        overrides["target"] = joined
    else:
        _notify(
            context,
            "targeted.cancel",
            "Targeted scrub aborted (no versions provided).",
            level="warning",
        )
        print("No target versions entered; aborting targeted scrub.")
        return
    if includes_raw:
        overrides["include"] = includes_raw
    _notify(
        context,
        "targeted.start",
        "Initiating targeted scrub run.",
        targets=overrides.get("target"),
        include=overrides.get("include"),
    )
    _plan_and_execute(context, overrides, label="targeted scrub")


def _menu_cleanup(context: MutableMapping[str, object]) -> None:
    _plan_and_execute(context, {"mode": "cleanup-only", "cleanup_only": True}, label="cleanup-only")


def _menu_diagnostics(context: MutableMapping[str, object]) -> None:
    _notify(context, "diagnostics.start", "Generating diagnostics plan from CLI menu.")
    plan_steps = _ensure_plan(context, {"mode": "diagnose", "diagnose": True})
    executor = cast(
        Callable[[list[dict[str, object]], Mapping[str, object] | None], bool | None],
        context["executor"],
    )
    inventory = context.get("inventory")
    executor(
        plan_steps,
        {"mode": "diagnose", "diagnose": True, "inventory": inventory},
    )
    _notify(context, "diagnostics.complete", "Diagnostics artifacts generated.")
    print("Diagnostics captured; no actions executed.")


def _menu_settings(context: MutableMapping[str, object]) -> None:
    args = context.get("args")
    input_func = cast(Callable[[str], str], context.get("input", input))
    if args is None:
        print("Settings unavailable (no argument namespace detected).")
        return

    while True:
        print("Current settings:")
        print(f"  1. Dry-run: {bool(getattr(args, 'dry_run', False))}")
        print(f"  2. Create restore point: {not bool(getattr(args, 'no_restore_point', False))}")
        print(f"  3. License cleanup enabled: {not bool(getattr(args, 'no_license', False))}")
        print(f"  4. Keep user templates: {bool(getattr(args, 'keep_templates', False))}")
        print(f"  5. Log directory: {getattr(args, 'logdir', '(default)')}")
        print(f"  6. Backup directory: {getattr(args, 'backup', '(disabled)')}")
        timeout_val = getattr(args, "timeout", None)
        print(f"  7. Timeout (seconds): {timeout_val if timeout_val is not None else '(default)'}")
        print("  8. Return to main menu")
        selection = input_func("Choose a setting to modify (1-8): ").strip()
        if selection == "1":
            value = not bool(getattr(args, "dry_run", False))
            args.dry_run = value
            _notify(context, "settings.dry_run", f"Dry-run set to {value}.", value=value)
        elif selection == "2":
            current = not bool(getattr(args, "no_restore_point", False))
            new_value = not current
            args.no_restore_point = not new_value
            _notify(
                context,
                "settings.restore_point",
                f"Create restore point set to {new_value}.",
                value=new_value,
            )
        elif selection == "3":
            current = not bool(getattr(args, "no_license", False))
            new_value = not current
            args.no_license = not new_value
            _notify(
                context,
                "settings.license",
                f"License cleanup enabled: {new_value}.",
                value=new_value,
            )
        elif selection == "4":
            value = not bool(getattr(args, "keep_templates", False))
            args.keep_templates = value
            _notify(context, "settings.templates", f"Keep templates set to {value}.", value=value)
        elif selection == "5":
            path_value = input_func("Enter log directory (blank for default): ").strip()
            args.logdir = path_value or None
            _notify(
                context,
                "settings.logdir",
                f"Log directory set to {path_value or '(default)'}.",
                value=path_value or None,
            )
        elif selection == "6":
            path_value = input_func("Enter backup directory (blank to disable): ").strip()
            args.backup = path_value or None
            _notify(
                context,
                "settings.backup",
                f"Backup directory set to {path_value or '(disabled)'}.",
                value=path_value or None,
            )
        elif selection == "7":
            raw_timeout = input_func("Enter timeout seconds (blank for default): ").strip()
            if raw_timeout:
                try:
                    timeout_val = int(raw_timeout)
                except ValueError:
                    _notify(
                        context,
                        "settings.timeout_invalid",
                        f"Timeout value {raw_timeout!r} is not an integer.",
                        level="warning",
                    )
                    print("Timeout must be an integer number of seconds.")
                    continue
                args.timeout = timeout_val
                _notify(
                    context,
                    "settings.timeout",
                    f"Timeout set to {timeout_val} seconds.",
                    value=timeout_val,
                )
            else:
                args.timeout = None
                _notify(context, "settings.timeout", "Timeout reset to default.", value=None)
        elif selection == "8":
            _notify(context, "settings.exit", "Returning to main menu from settings.")
            break
        else:
            _notify(
                context,
                "settings.invalid",
                f"Invalid settings selection {selection!r}.",
                level="warning",
            )
            print("Please choose a valid option (1-8).")


def _menu_exit(context: MutableMapping[str, object]) -> None:
    context["running"] = False
    _notify(context, "ui.exit", "Exiting Office Janitor interactive CLI.")
    print("Exiting Office Janitor.")


def _plan_and_execute(
    context: MutableMapping[str, object], overrides: Mapping[str, object], *, label: str
) -> None:
    _notify(context, "plan.start", f"Planning run for {label} mode.", overrides=dict(overrides))
    plan_steps = _ensure_plan(context, overrides)
    executor = cast(
        Callable[[list[dict[str, object]], Mapping[str, object] | None], bool | None],
        context["executor"],
    )
    payload = dict(overrides)
    if "inventory" not in payload and context.get("inventory") is not None:
        payload["inventory"] = context.get("inventory")
    summary = plan_module.summarize_plan(plan_steps)
    _notify(
        context,
        "plan.ready",
        f"Plan ready for {label} mode with {summary.get('total_steps', 0)} steps.",
        summary=summary,
    )
    confirm_helper = context.get("confirm")
    args = context.get("args")
    dry_run = bool(getattr(args, "dry_run", False)) if args is not None else False
    force = bool(getattr(args, "force", False)) if args is not None else False
    input_func = cast(Callable[[str], str], context.get("input", input))
    if callable(confirm_helper):
        proceed = confirm_helper(
            dry_run=dry_run,
            force=force,
            input_func=input_func,
            interactive=True,
        )
        if not proceed:
            _notify(
                context,
                "execution.cancelled",
                "Scrub cancelled by user confirmation prompt.",
                level="warning",
            )
            print("Scrub cancelled.")
            return
        payload["confirmed"] = True
    if "force" not in payload:
        payload["force"] = force
    payload["input_func"] = input_func
    payload["interactive"] = True
    executor(plan_steps, payload)
    _notify(context, "execution.complete", f"Execution finished for {label} mode.")


def _ensure_plan(
    context: MutableMapping[str, object], overrides: Mapping[str, object]
) -> list[dict[str, object]]:
    planner = cast(
        Callable[[Mapping[str, object], Mapping[str, object] | None], list[dict[str, object]]],
        context["planner"],
    )
    inventory = context.get("inventory")
    if inventory is None:
        detector = cast(Callable[[], Mapping[str, object]], context["detector"])
        _notify(context, "detect.lazy", "Collecting inventory prior to planning.")
        inventory = detector()
        context["inventory"] = inventory
    plan_steps = planner(inventory, overrides)
    context["plan"] = plan_steps
    print(f"Plan contains {len(plan_steps)} step(s).")
    return plan_steps


def _summarize_inventory(inventory: Mapping[str, object]) -> dict[str, int]:
    summary: dict[str, int] = {}
    for key, items in inventory.items():
        count = _count_items(items)
        summary[str(key)] = count
    return summary


def _notify(
    context: Mapping[str, object],
    event: str,
    message: str,
    *,
    level: str = "info",
    **payload: object,
) -> None:
    """!
    @brief Emit user-facing and structured log updates for menu actions.
    """

    human_logger = context.get("human_logger")
    if isinstance(human_logger, logging.Logger):
        log_func = getattr(human_logger, level, human_logger.info)
        log_func(message)

    machine_logger = context.get("machine_logger")
    if isinstance(machine_logger, logging.Logger):
        extra: dict[str, object] = {"event": "ui_progress", "name": event}
        if message:
            extra["message"] = message
        if payload:
            extra["data"] = dict(payload)
        machine_logger.info("ui_progress", extra=extra)

    record = {"event": event, "message": message}
    if payload:
        record["data"] = dict(payload)

    emitter = context.get("emit_event")
    if callable(emitter):
        emitter(event, message=message, **payload)
    else:
        queue = context.get("event_queue")
        append = getattr(queue, "append", None)
        if callable(append):
            append(record)


def _count_items(items: object) -> int:
    if isinstance(items, Mapping):
        return len(items)
    if isinstance(items, MutableMapping):
        return len(items)
    if isinstance(items, (list, tuple, set, frozenset)):
        return len(items)
    if isinstance(items, str):
        return 1 if items else 0
    if isinstance(items, Iterable):
        return sum(1 for _ in items)
    return 0
