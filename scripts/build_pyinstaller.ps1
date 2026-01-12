<#!
.SYNOPSIS
Build the Office Janitor PyInstaller onefile executable and prepare release artifacts.

.DESCRIPTION
Invokes PyInstaller with the configuration described in spec.md, ensures the src
package directory is available on PYTHONPATH, and archives the generated
executable into the artifacts folder for distribution.
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

$command = "pyinstaller --clean --onefile --uac-admin --name office-janitor src/office_janitor/main.py --paths src"

Write-Host "Running PyInstaller: $command"
Invoke-Expression $command | Write-Output

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
