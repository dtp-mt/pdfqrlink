# QRPdfAnnotator.spec
# pyinstaller QRPdfAnnotator.spec でビルド
from PyInstaller.utils.hooks import collect_all
import sys
block_cipher = None

# tkinterdnd2 のデータ一式を収集
datas, binaries, hiddenimports = collect_all('tkinterdnd2')

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports + ['tkinter', 'cv2', 'fitz'],
    hookspath=[],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz, a.scripts, a.binaries, a.zipfiles, a.datas,
    name='QRPdfAnnotator',
    console=False,  # windowed
    # target_arch='universal2',  # ← 本当に狙うなら。依存がすべて universal2 必須。
)

app = BUNDLE(
    exe,
    name='QRPdfAnnotator.app',
    icon=None,
    bundle_identifier='com.example.qrpdfannotator'
)
