# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# tkinterdnd2 のリソース＆サブモジュールを収集
_tkdnd_datas = collect_data_files('tkinterdnd2')
_tkdnd_hidden = collect_submodules('tkinterdnd2')

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=_tkdnd_datas,           # ← 追加
    hiddenimports=_tkdnd_hidden,  # ← 追加
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name='PDF内QRコードに注釈追加',
    debug=False, strip=False, upx=False,
    console=False,
)
coll = COLLECT(
    exe, a.binaries, a.datas,
    strip=False, upx=False, upx_exclude=[],
    name='PDF内QRコードに注釈追加'
)
app = BUNDLE(
    coll,
    name='PDF内QRコードに注釈追加.app',
    icon='a0zmx-dwaac.icns',
    bundle_identifier='jp.example.qr-pdf',
    info_plist={
        'CFBundleName': 'PDF内QRコードに注釈追加',
        'CFBundleShortVersionString': '1.0.0',
        'CFBundleVersion': '1.0.0',
        'NSHighResolutionCapable': 'True',
    }
)
