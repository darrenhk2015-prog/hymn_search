# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules
from PyInstaller.utils.hooks import collect_all

datas = [('assets', 'assets'), ('hymn_remote\\static', 'hymn_remote\\static'), ('hymn_remote\\data', 'hymn_remote\\data')]
binaries = []
hiddenimports = ['fitz', 'version', 'hymn_features.fuzzy', 'hymn_features.session_store', 'hymn_features.preview', 'hymn_features.mixin', 'hymn_features.voice_input', 'qrcode', 'qrcode.image', 'qrcode.image.pil', 'qrcode.main', 'PIL', 'PIL.Image', 'pypinyin', 'speech_recognition', 'sounddevice', 'numpy', 'audioop', 'audioop_lts', 'hymn_remote.api', 'hymn_remote.web_hymn_map', 'hymn_remote.gdrive_hymn_map', 'hymn_remote.gdrive_index_build', 'hymn_remote.viewer_hymn_map', 'hymn_features.gdrive_map_dialog', 'gdown', 'hymn_remote.server', 'hymn_remote.tunnel', 'uvicorn.logging', 'uvicorn.loops', 'uvicorn.loops.auto', 'uvicorn.protocols.http.auto', 'uvicorn.lifespan.on']
hiddenimports += collect_submodules('qrcode')
hiddenimports += collect_submodules('uvicorn')
hiddenimports += collect_submodules('fastapi')
hiddenimports += collect_submodules('starlette')
tmp_ret = collect_all('pillow')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
try:
    sd_ret = collect_all('sounddevice')
    datas += sd_ret[0]; binaries += sd_ret[1]; hiddenimports += sd_ret[2]
except Exception:
    pass


a = Analysis(
    ['hymn_search.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='HymnSearch',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['assets\\app.ico'],
)
