# office-janitor

A scrubber that helps you fully remove Office easily

## Development tooling

The project is configured for a ``src`` layout and ships without runtime dependencies.
Install the optional developer extras to enable formatting, linting, type-checking,
testing, and packaging helpers:

```bash
python -m pip install --upgrade pip
python -m pip install .[dev]
```

### Formatting

Run Black in check-only mode to validate formatting prior to submitting a change:

```bash
black --check .
```

### Linting and static analysis

Ruff and MyPy provide linting and type coverage for the package source:

```bash
ruff check .
mypy src
```

### Tests

Execute the test suite with Pytest. CI runs this matrix on Windows for Python 3.9
and 3.11.

```bash
pytest
```

### Building the PyInstaller artifact

The PyInstaller command wired into CI mirrors the specification:

```powershell
pyinstaller --clean --onefile --uac-admin --name OfficeJanitor office_janitor.py --paths src
```

The repository also includes a helper script that installs PyInstaller, performs
the build, and archives the resulting executable and manifest into ``artifacts``:

```powershell
pwsh scripts/build_pyinstaller.ps1
```

Artifacts are written to ``dist/`` by PyInstaller and zipped to
``artifacts/OfficeJanitor-win64.zip`` for distribution.
