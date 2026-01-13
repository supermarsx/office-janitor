<#!
.SYNOPSIS
Build the Office Janitor PyInstaller onefile executable and prepare release artifacts.

.DESCRIPTION
Invokes PyInstaller using the pre-configured office-janitor.spec file which includes
all OEM files (XML configs, setup.exe, OfficeClickToRun.exe) and the VERSION file.
Archives the generated executable into the artifacts folder for distribution.
#>

param(
    [string]$Python = "python",
    [string]$DistFolder = "dist",
    [string]$ArtifactFolder = "artifacts"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

Write-Host "Installing PyInstaller..."
& $Python -m pip install --upgrade pyinstaller | Write-Output

$specFile = Join-Path $repoRoot "office-janitor.spec"
if (-not (Test-Path $specFile)) {
    throw "Spec file not found at $specFile"
}

Write-Host "Building with spec file: $specFile"
& $Python -m PyInstaller --clean --noconfirm $specFile | Write-Output

if (-not (Test-Path $ArtifactFolder)) {
    New-Item -ItemType Directory -Path $ArtifactFolder | Out-Null
}

$executablePath = Join-Path $repoRoot "$DistFolder/office-janitor.exe"
if (-not (Test-Path $executablePath)) {
    throw "Expected executable not found at $executablePath"
}

$manifestPath = Join-Path $repoRoot "$DistFolder/office-janitor.exe.manifest"
$artifactZip = Join-Path $ArtifactFolder "office-janitor-win64.zip"

$itemsToArchive = @($executablePath)
if (Test-Path $manifestPath) {
    $itemsToArchive += $manifestPath
}

Write-Host "Creating artifact archive at $artifactZip"
Compress-Archive -Path $itemsToArchive -DestinationPath $artifactZip -Force

Write-Host "Artifacts available in $ArtifactFolder"
