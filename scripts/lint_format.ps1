<#!
.SYNOPSIS
Check or fix formatting and lint issues using Black and Ruff.

.DESCRIPTION
Runs Ruff followed by Black with a Windows-friendly default configuration.
Use -Fix to apply changes in place; otherwise the tools run in check-only
mode for CI compatibility.
#>
param(
    [string]$Python = "python",
    [switch]$Fix,
    [string[]]$Paths = @("src", "tests", "office_janitor.py", "scripts")
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

Write-Host "Running Ruff..." -ForegroundColor Cyan
$ruffArgs = @("check") + $Paths
if ($Fix) {
    $ruffArgs += "--fix"
}
& $Python -m ruff @ruffArgs

Write-Host "Running Black..." -ForegroundColor Cyan
$blackArgs = @("--check") + $Paths
if ($Fix) {
    $blackArgs = $Paths
}
& $Python -m black @blackArgs
