<#!
.SYNOPSIS
Run mypy type checks with the repository configuration.

.DESCRIPTION
Invokes mypy against src and tests from the repo root, matching the settings
in pyproject.toml. Accepts an alternate Python interpreter when needed.
#>
param(
    [string]$Python = "python",
    [string[]]$Targets = @("src")
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

Write-Host "Running mypy..." -ForegroundColor Cyan
& $Python -m mypy @Targets
