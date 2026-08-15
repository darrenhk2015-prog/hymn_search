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

$VenvPython = Join-Path $Root '.venv\Scripts\python.exe'
if (Test-Path $VenvPython) {
    $Python = $VenvPython
    Write-Host "Using venv: $Python"
} else {
    $Python = 'python'
    Write-Host "Using system python (no .venv found)"
}

Write-Host "Installing dependencies..."
& $Python -m pip install -r requirements.txt -q
& $Python -c "import qrcode; from PIL import Image; print('qrcode + Pillow OK')"

if (-not (Test-Path 'assets\app.ico')) {
    & $Python scripts\make_icon.py
}

$GdriveIndex = Join-Path $Root 'hymn_remote\data\gdrive_hymn_index.json'
if (-not (Test-Path $GdriveIndex)) {
    Write-Host "Building Google Drive hymn index (first time)..."
    & $Python -m pip install gdown -q
    & $Python scripts\build_gdrive_hymn_map.py
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
    $buildDate = Get-Date -Format 'yyyy-MM-dd'
    if ($versionContent -match 'BUILD_DATE\s*=') {
        $versionContent = [regex]::Replace(
            $versionContent,
            "BUILD_DATE\s*=\s*'[^']*'",
            "BUILD_DATE = '$buildDate'"
        )
    } else {
        $versionContent = $versionContent.TrimEnd() + "`nBUILD_DATE = '$buildDate'`n"
    }
    [System.IO.File]::WriteAllText($VersionPath, $versionContent)
    Write-Host "Build version: $buildVer"
} else {
    throw "Cannot parse APP_VERSION in version.py"
}

& $Python -m PyInstaller --noconfirm --onefile --windowed `
    --name "HymnSearch" `
    --icon "assets\app.ico" `
    --add-data "assets;assets" `
    --add-data "hymn_remote\static;hymn_remote\static" `
    --add-data "hymn_remote\data;hymn_remote\data" `
    --hidden-import fitz `
    --hidden-import version `
    --hidden-import hymn_features.fuzzy `
    --hidden-import hymn_features.session_store `
    --hidden-import hymn_features.preview `
    --hidden-import hymn_features.mixin `
    --hidden-import qrcode `
    --hidden-import qrcode.image `
    --hidden-import qrcode.image.pil `
    --hidden-import qrcode.main `
    --hidden-import PIL `
    --hidden-import PIL.Image `
    --collect-submodules qrcode `
    --collect-all pillow `
    --hidden-import pypinyin `
    --hidden-import hymn_remote.api `
    --hidden-import hymn_remote.web_hymn_map `
    --hidden-import hymn_remote.gdrive_hymn_map `
    --hidden-import hymn_remote.gdrive_index_build `
    --hidden-import hymn_remote.viewer_hymn_map `
    --hidden-import hymn_features.gdrive_map_dialog `
    --hidden-import gdown `
    --hidden-import hymn_remote.server `
    --hidden-import hymn_remote.tunnel `
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
