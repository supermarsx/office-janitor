# TODO

- [ ] Mypy compliance (currently ~151 errors):
  - [ ] `plan.py`: pass mutable sequences into `_augment_auto_all_c2r_inventory`, remove unused ignores, replace `int/list/dict` casts on `object` with validated conversions.
  - [ ] UI/TUI: add typed app_state/ logger references, type event queues (`deque[dict[str, object]]`), clean unused ignores, and ensure `msvcrt` stub typing is acceptable.
  - [ ] Scrub/Plan orchestration: fix `build_plan` input types (`dict[str, Sequence[dict]]`), scrub result typing, avoid `dict(obj)`/`list(obj)` on unknowns, ensure `c2r_uninstall`/`msi_uninstall` interfaces are typed.
  - [ ] Safety/registry/fs tools: remove unused `type: ignore` comments; ensure winreg stubs cover attr-defined errors; add explicit `Mapping` types where `.get` is used.
  - [ ] Detect/licensing/off_scrub: validate collections before casting to `list`/`dict`; resolve duplicate variable names and assignment types.
  - [ ] Logging: `_SizedTimedRotatingFileHandler` args typed; remove unused ignores.
- [ ] CI workflows: split monolithic `.github/workflows/ci.yml` into `format.yml`, `lint.yml`, `test.yml`, `build.yml`, `publish-pypi.yml`, and `release.yml` per spec.
