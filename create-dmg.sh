#!/usr/bin/env bash
set -euo pipefail

APP_BUNDLE="${1:-dist/PDF内QRコードに注釈追加.app}"
DMG_NAME="${2:-PDF内QRコードに注釈追加.dmg}"
VOLNAME="${3:-PDF内QRコードに注釈追加}"

if [ ! -d "${APP_BUNDLE}" ]; then
  echo "ERROR: App bundle not found: ${APP_BUNDLE}" >&2
  ls -la dist || true
  exit 1
fi

rm -f "${DMG_NAME}"
mkdir -p dmg_contents
cp -R "${APP_BUNDLE}" dmg_contents/

create-dmg \
  --volname "${VOLNAME}" \
  --window-pos 200 120 \
  --window-size 600 300 \
  --icon-size 100 \
  --app-drop-link 480 120 \
  "${DMG_NAME}" \
  dmg_contents

rm -rf dmg_contents
