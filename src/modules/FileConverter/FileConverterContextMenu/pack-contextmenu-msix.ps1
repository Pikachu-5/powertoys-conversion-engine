param(
    [string]$Platform = "x64",
    [string]$Configuration = "Debug",
    [int]$KeepRecent = 5
)

$ErrorActionPreference = "Stop"

function Get-MakeAppxPath {
    $command = Get-Command MakeAppx.exe -ErrorAction SilentlyContinue
    if ($null -ne $command) {
        return $command.Source
    }

    $kitsRoot = "C:\Program Files (x86)\Windows Kits\10\bin"
    if (-not (Test-Path $kitsRoot)) {
        throw "MakeAppx.exe was not found in PATH and Windows Kits bin folder does not exist."
    }

    $candidates = Get-ChildItem -Path $kitsRoot -Recurse -Filter MakeAppx.exe -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -match "\\x64\\MakeAppx\.exe$" } |
        Sort-Object FullName -Descending

    if ($null -eq $candidates -or $candidates.Count -eq 0) {
        throw "MakeAppx.exe was not found under $kitsRoot."
    }

    return $candidates[0].FullName
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = (Resolve-Path (Join-Path $scriptDir "..\..\..\..")).Path
$outDir = Join-Path $repoRoot "$Platform\$Configuration\WinUI3Apps"

New-Item -Path $outDir -ItemType Directory -Force | Out-Null

$contextMenuDll = Join-Path $outDir "PowerToys.FileConverterContextMenu.dll"
if (-not (Test-Path $contextMenuDll)) {
    throw "Context menu DLL was not found at $contextMenuDll. Build FileConverterContextMenu first."
}

$stagingRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("FileConverterContextMenuPackage_" + [System.Guid]::NewGuid().ToString("N"))
New-Item -Path $stagingRoot -ItemType Directory -Force | Out-Null

try {
    Copy-Item -Path (Join-Path $scriptDir "AppxManifest.xml") -Destination (Join-Path $stagingRoot "AppxManifest.xml") -Force
    Copy-Item -Path (Join-Path $scriptDir "Assets") -Destination (Join-Path $stagingRoot "Assets") -Recurse -Force
    Copy-Item -Path $contextMenuDll -Destination (Join-Path $stagingRoot "PowerToys.FileConverterContextMenu.dll") -Force

    $makeAppx = Get-MakeAppxPath
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $timestampedPackage = Join-Path $outDir "FileConverterContextMenuPackage.$timestamp.msix"
    $stablePackage = Join-Path $outDir "FileConverterContextMenuPackage.msix"

    Write-Host "Using MakeAppx:" $makeAppx
    Write-Host "Packaging to:" $timestampedPackage

    & $makeAppx pack /d $stagingRoot /p $timestampedPackage /nv
    if ($LASTEXITCODE -ne 0) {
        throw "MakeAppx packaging failed with exit code $LASTEXITCODE."
    }

    try {
        Copy-Item -Path $timestampedPackage -Destination $stablePackage -Force -ErrorAction Stop
        Write-Host "Updated stable package:" $stablePackage
    }
    catch {
        Write-Warning "Stable package copy skipped because destination appears locked."
    }

    $allPackages = Get-ChildItem -Path $outDir -Filter "FileConverterContextMenuPackage.*.msix" |
        Sort-Object LastWriteTime -Descending

    if ($allPackages.Count -gt $KeepRecent) {
        $allPackages | Select-Object -Skip $KeepRecent | Remove-Item -Force -ErrorAction SilentlyContinue
    }

    Write-Host "Timestamped package ready:" $timestampedPackage
}
finally {
    Remove-Item -Path $stagingRoot -Recurse -Force -ErrorAction SilentlyContinue
}
