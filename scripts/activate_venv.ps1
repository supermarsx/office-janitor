<#! 
.SYNOPSIS
Create (if needed) and activate the repo's local Python virtual environment.

.DESCRIPTION
Bootstraps a `.venv` in the repository root using the chosen Python interpreter,
optionally installs development dependencies, and then launches a PowerShell
session with the environment activated. If you prefer to stay in the current
session, you can dot-source this script instead: `. .\scripts\activate_venv.ps1 -NoShell`.
#>
[CmdletBinding()]
param(
    [string]$Python = "python",
    [switch]$InstallDev,
    [string]$Extras = "dev",
    [switch]$NoShell
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

$venvPath = Join-Path $repoRoot ".venv"
$activatePath = Join-Path $venvPath "Scripts\Activate.ps1"

if (-not (Test-Path $venvPath)) {
    Write-Host "Creating virtual environment at $venvPath..." -ForegroundColor Cyan
    & $Python -m venv $venvPath
}

Write-Host "Upgrading pip and wheel in the venv..." -ForegroundColor Cyan
& "$venvPath\Scripts\python.exe" -m pip install --upgrade pip wheel

if ($InstallDev) {
    $extrasSpec = if ($Extras) { "[${Extras}]" } else { "" }
    Write-Host "Installing editable package with extras spec '$extrasSpec'..." -ForegroundColor Cyan
    & "$venvPath\Scripts\python.exe" -m pip install -e ".${extrasSpec}"
}

if ($NoShell) {
    Write-Host "Environment prepared. Activate with:`n  . $activatePath" -ForegroundColor Green
    return
}

Write-Host "Launching a new PowerShell session with the venv activated..." -ForegroundColor Green
powershell.exe -NoExit -Command "& { Set-Location '$repoRoot'; . '$activatePath' }"
