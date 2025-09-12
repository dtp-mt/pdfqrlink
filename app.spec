name: macOS ARM .app build

on:
  workflow_dispatch:

jobs:
  build:
    runs-on: macos-14
    steps:
      - uses: actions/checkout@v4

      - name: Setup Python
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: Install dependencies
        run: |
          python -m pip install -U pip wheel setuptools
          pip install pymupdf Pillow numpy opencv-python-headless zxing-cpp tkinterdnd2-universal pyinstaller

      - name: Build .app with PyInstaller
        run: |
          # tkinterdnd2 のデータ同梱を強めに（tkdnd2.9配下も確実に拾う）
          cat > hook-tkinterdnd2.py << 'PY'
          from PyInstaller.utils.hooks import collect_data_files
          datas = collect_data_files('tkinterdnd2', includes=['tkdnd2.9/*','TKDND*','*.txt','*.md'])
          PY
          pyinstaller -y app.spec

      # --- ここから署名 & 公証 追加 ---
      - name: Import signing certificate
        env:
          P12_BASE64: ${{ secrets.MAC_CERT_P12_BASE64 }}
          P12_PASSWORD: ${{ secrets.MAC_CERT_P12_PASSWORD }}
        run: |
          echo "$P12_BASE64" | base64 --decode > signing.p12
          KEYCHAIN=build.keychain
          security create-keychain -p "" $KEYCHAIN
          security default-keychain -s $KEYCHAIN
          security unlock-keychain -p "" $KEYCHAIN
          security import signing.p12 -k $KEYCHAIN -P "$P12_PASSWORD" -T /usr/bin/codesign
          # codesign が非対話で使えるように権限付与
          security set-key-partition-list -S apple-tool:,apple:,codesign: -s -k "" $KEYCHAIN
          rm -f signing.p12

      - name: Codesign .app (deep + hardened)
        env:
          CODE_SIGN_IDENTITY: ${{ secrets.MAC_CERT_IDENTITY }}
        run: |
          APP="dist/PDF内QRコードに注釈追加.app"
          # 念のため、ネストした dylib / so / 実行ファイルを先に署名
          find "$APP/Contents" -type f \( -name "*.dylib" -o -name "*.so" -o -perm -111 \) -print0 | \
            xargs -0 -I{} codesign --force --options runtime --timestamp --sign "$CODE_SIGN_IDENTITY" "{}"
          # バンドル本体を署名
          codesign --deep --force --options runtime --timestamp --sign "$CODE_SIGN_IDENTITY" "$APP"
          codesign --verify --deep --strict --verbose=2 "$APP"

      - name: Notarize .app (zip submit -> wait -> staple)
        env:
          APPLE_ID: ${{ secrets.APPLE_ID }}
          APPLE_TEAM_ID: ${{ secrets.APPLE_TEAM_ID }}
          APPLE_APP_PASS: ${{ secrets.APPLE_APP_PASS }}
        run: |
          APP="dist/PDF内QRコードに注釈追加.app"
          /usr/bin/zip -ry "App.zip" "$APP"
          # Apple ID + App-specific password で notarytool 認証
          xcrun notarytool submit "App.zip" \
            --apple-id "$APPLE_ID" \
            --team-id "$APPLE_TEAM_ID" \
            --password "$APPLE_APP_PASS" \
            --wait
          # 公証結果を .app に貼り付け
          xcrun stapler staple "$APP"
          # 念のため検証
          spctl -a -vvvv "$APP"
          codesign --verify --deep --strict --verbose=2 "$APP"

      - name: Upload notarized .app bundle
        uses: actions/upload-artifact@v4
        with:
          name: qr-pdf-annotator-arm
          path: 'dist/PDF内QRコードに注釈追加.app'
      # --- 署名 & 公証 ここまで ---

      - name: Create notarized .dmg from stapled .app
        run: |
          # 依存を減らすため mac 純正で DMG 作成（brew なし）
          APP="dist/PDF内QRコードに注釈追加.app"
          DMG="PDF内QRコードに注釈追加.dmg"
          rm -f "$DMG"
          hdiutil create -volname "PDF内QRコードに注釈追加" -srcfolder "$APP" -ov -format UDZO "$DMG"
          # DMG 自体にも staple（任意だがおすすめ）
          xcrun stapler staple "$DMG" || true

      - name: Upload .dmg
        uses: actions/upload-artifact@v4
        with:
          name: qr-pdf-annotator-dmg
          path: 'PDF内QRコードに注釈追加.dmg'
