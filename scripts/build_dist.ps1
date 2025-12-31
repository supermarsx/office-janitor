<#!
.SYNOPSIS
Build source and wheel distributions using python -m build.

.DESCRIPTION
Creates sdist and wheel artifacts into the specified output directory. Use
-RefreshTools to ensure the build module is present/updated on Windows hosts.
#>
param(
    [string]$Python = "python",
    [string]$Output = "dist",
    [switch]$RefreshTools
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

if ($RefreshTools) {
    Write-Host "Ensuring build module is available..." -ForegroundColor Cyan
    & $Python -m pip install --upgrade build
}

Write-Host "Building distributions into '$Output'..." -ForegroundColor Cyan
& $Python -m build --outdir $Output
