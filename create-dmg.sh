
#!/bin/bash
set -e

APP_NAME="QR PDF Annotator"
APP_BUNDLE="dist/$APP_NAME.app"
DMG_NAME="$APP_NAME.dmg"

# Create temporary folder for DMG contents
mkdir -p dmg_contents
cp -R "$APP_BUNDLE" dmg_contents/

# Create DMG using create-dmg (must be installed via brew)
create-dmg   --volname "$APP_NAME"   --window-pos 200 120   --window-size 600 300   --icon-size 100   --app-drop-link 480 120   "$DMG_NAME"   dmg_contents

# Cleanup
rm -rf dmg_contents
