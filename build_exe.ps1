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

pyinstaller --noconfirm --onefile --windowed `
    --name "HymnSearch" `
    --icon "assets\app.ico" `
    --add-data "assets;assets" `
    --hidden-import fitz `
    hymn_search.py

Write-Host ""
Write-Host "Done: dist\HymnSearch.exe"
