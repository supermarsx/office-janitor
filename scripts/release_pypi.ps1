<#
.SYNOPSIS
Automatically bump 0.0.x, build distributions, and publish to PyPI.

.DESCRIPTION
Calls scripts/release_pypi.py which determines the next available 0.0.x
version by checking PyPI, writes src/office_janitor/VERSION, builds wheel and
sdist, and uploads with Twine.
#>

param(
    [string]$Python = "python",
    [switch]$NoPublish,
    [switch]$DryRun,
    [switch]$TestPyPI,
    [switch]$RefreshTools,
    [int]$MaxAttempts = 5
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

$args = @("scripts/release_pypi.py", "--python", $Python, "--max-attempts", "$MaxAttempts")

if (-not $NoPublish) {
    $args += "--publish"
}
if ($DryRun) {
    $args += "--dry-run"
}
if ($TestPyPI) {
    $args += "--repository-url"
    $args += "https://test.pypi.org/legacy/"
}
if ($RefreshTools) {
    $args += "--refresh-tools"
}

Write-Host "Running PyPI release automation..." -ForegroundColor Cyan
& $Python @args
