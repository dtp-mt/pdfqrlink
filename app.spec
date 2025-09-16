# QRPdfAnnotator.spec
# ビルド: pyinstaller --clean -y QRPdfAnnotator.spec

from PyInstaller.utils.hooks import (
    collect_all, collect_data_files, collect_submodules, collect_dynamic_libs
)
import sys

block_cipher = None

# --- tkinterdnd2 一式を収集（バイナリ含む） ---
datas, binaries, hiddenimports = collect_all('tkinterdnd2')  # ✅
# 参考: tkinterdnd2 は PyInstaller 用フックを推奨。collect_all でも可
# https://pypi.org/project/tkinterdnd2/ 参照

# --- ttkbootstrap のデータ( themes.json など ) を同梱 ---
datas += collect_data_files('ttkbootstrap')  # ✅ 推奨
# 参考: themes.json 見つからずで落ちる既知事例
# https://stackoverflow.com/questions/67850998/ttkbootstrap-not-working-with-pyinstaller

# --- OpenCV 取りこぼし対策（任意だが安定） ---
# サブモジュールを明示収集
hiddenimports += collect_submodules('cv2')
# ネイティブライブラリを必要に応じ追加（PyInstaller 6 以降）
try:
    binaries += collect_dynamic_libs('cv2')
except Exception:
    pass
# 参考: cv2 ローダの挙動で PyInstaller が苦戦するケースあり
# https://github.com/pyinstaller/pyinstaller/issues/6889
# https://pyinstaller.org/en/stable/hooks.html

# --- (任意) PyMuPDF の追加データ ---
# datas += collect_data_files('fitz')  # 必要時のみ
# 参考: 'fitz' は PyMuPDF のレガシー名。pypi の別 'fitz' と混同注意
# https://pymupdf.qubitpi.org/en/latest/installation.html

a = Analysis(
    ['app.py'],
    pathex=['.'],          # ルートを明示
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports + ['tkinter', 'fitz'],  # 'fitz' を使っているため残す
    hookspath=[],
    noarchive=False,       # そのままでOK（サイズ・起動速度のバランス）
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz, a.scripts, a.binaries, a.zipfiles, a.datas,
    name='QRPdfAnnotator',
    console=False,  # windowed
    # target_arch='universal2',  # ← Universal2 を本気で狙うなら、Python本体/拡張含め全依存が universal2 必須
)

app = BUNDLE(
    exe,
    name='QRPdfAnnotator.app',
    icon=None,  # あるなら .icns を指定可
    bundle_identifier='com.example.qrpdfannotator'
)
