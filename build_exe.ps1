# Build standalone Windows exe for 詩歌冊搜索
# Usage (PowerShell):
#   cd C:\python\hymn_search
#   .\.venv\Scripts\Activate.ps1
#   pip install -r requirements.txt
#   python scripts\make_icon.py
#   .\build_exe.ps1

$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

if (-not (Test-Path 'assets\app.ico')) {
    python scripts\make_icon.py
}

$VersionPath = Join-Path $Root 'version.py'
if (-not (Test-Path $VersionPath)) {
    @'
"""Application build version (auto-incremented by build_exe.ps1)."""
APP_VERSION = 0
'@ | Set-Content -Path $VersionPath -Encoding UTF8
}

$versionContent = Get-Content $VersionPath -Raw
if ($versionContent -match 'APP_VERSION\s*=\s*(\d+)') {
    $buildVer = [int]$Matches[1] + 1
    $versionContent = [regex]::Replace(
        $versionContent,
        'APP_VERSION\s*=\s*\d+',
        "APP_VERSION = $buildVer"
    )
    [System.IO.File]::WriteAllText($VersionPath, $versionContent)
    Write-Host "Build version: $buildVer"
} else {
    throw "Cannot parse APP_VERSION in version.py"
}

pyinstaller --noconfirm --onefile --windowed `
    --name "HymnSearch" `
    --icon "assets\app.ico" `
    --add-data "assets;assets" `
    --add-data "hymn_remote\static;hymn_remote\static" `
    --hidden-import fitz `
    --hidden-import version `
    --hidden-import hymn_remote.api `
    --hidden-import hymn_remote.server `
    --hidden-import uvicorn.logging `
    --hidden-import uvicorn.loops `
    --hidden-import uvicorn.loops.auto `
    --hidden-import uvicorn.protocols.http.auto `
    --hidden-import uvicorn.lifespan.on `
    --collect-submodules uvicorn `
    --collect-submodules fastapi `
    --collect-submodules starlette `
    hymn_search.py

Write-Host ""
Write-Host "Done: dist\HymnSearch.exe"
