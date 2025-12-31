<#!
.SYNOPSIS
Execute the pytest suite.

.DESCRIPTION
Runs pytest from the repository root. Extra arguments are forwarded directly
to pytest, making it easy to select subsets (`-k`) or enable verbose output.
#>
param(
    [string]$Python = "python",
    [string[]]$Args = @()
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

Write-Host "Running pytest..." -ForegroundColor Cyan
& $Python -m pytest @Args
