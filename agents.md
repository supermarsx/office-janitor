# office-janitor agent instructions

## Scope
These instructions apply to the entire repository unless a nested `agents.md` overrides them.

## Architectural expectations
- Keep the project aligned with the layout defined in `spec.md`, retaining `office_janitor.py` as the shim and using the `src/office_janitor/` package structure for implementation modules.
- Follow the responsibilities described in the spec for each module (e.g., detection logic in `detect.py`, uninstall orchestration in `scrub.py`). If a new capability is added, place it in the module that best matches the spec’s intent and update documentation/tests accordingly.

## Coding guidelines
- Target Python 3.9+ and rely on the standard library only.
- Preserve cross-version uninstall support for MSI and Click-to-Run Office releases. Add constants/data-driven mappings instead of hard-coding logic when possible.
- Guard destructive operations behind explicit flags as described in the spec (`--dry-run`, targeted scrubs, etc.).
- Maintain structured logging (human + JSONL) consistent with `logging_ext` expectations.

## Testing & quality
- When adding functionality, extend or create tests under `tests/` that reflect the scenarios outlined in the spec (detection, planning, safety, registry tools).
- Ensure changes stay compatible with the CI workflows described in `spec.md` (Black, Ruff, MyPy, Pytest, PyInstaller build).

## Pull requests
- Summaries should reference the spec-aligned features affected (e.g., detection, uninstall, licensing) and mention any test coverage updates.
