# Office Janitor Winget Manifests

This directory contains Windows Package Manager (winget) manifests for Office Janitor.

## Structure

```
winget/
├── supermarsx.office-janitor/
│   ├── latest/                    # Symlink to current version
│   │   ├── supermarsx.office-janitor.yaml
│   │   ├── supermarsx.office-janitor.installer.yaml
│   │   └── supermarsx.office-janitor.locale.en-US.yaml
│   └── YY.X/                      # Version-specific manifests
│       └── ...
```

## Installation

```powershell
# From winget-pkgs community repository (after PR merge)
winget install supermarsx.office-janitor

# Or from local manifest
winget install --manifest winget/supermarsx.office-janitor/latest/
```

## Submitting to winget-pkgs

To submit to the official winget-pkgs repository:

1. Fork [microsoft/winget-pkgs](https://github.com/microsoft/winget-pkgs)
2. Copy the version folder to `manifests/s/supermarsx/office-janitor/YY.X/`
3. Validate: `winget validate --manifest <path>`
4. Create a pull request

The CI pipeline automatically updates these manifests on each rolling release.
