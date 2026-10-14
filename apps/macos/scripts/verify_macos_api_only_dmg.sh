#!/bin/bash
set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
VERSION="$(node -p "require('$PROJECT_ROOT/package.json').version")"
DMG_PATH="${1:-$PROJECT_ROOT/src-tauri/target/release/bundle/dmg/VisualTeX_${VERSION}_aarch64-no-ocr.dmg}"

if [[ ! -f "$DMG_PATH" ]]; then
  echo "Missing API-only DMG: $DMG_PATH" >&2
  exit 1
fi

hdiutil verify "$DMG_PATH"
MOUNT_POINT="$(mktemp -d "${TMPDIR:-/tmp}/visualtex-api-only.XXXXXX")"
cleanup() {
  hdiutil detach "$MOUNT_POINT" -quiet >/dev/null 2>&1 || true
  rmdir "$MOUNT_POINT" >/dev/null 2>&1 || true
}
trap cleanup EXIT

hdiutil attach "$DMG_PATH" -readonly -nobrowse -mountpoint "$MOUNT_POINT" -quiet
APP="$MOUNT_POINT/VisualTeX.app"
RESOURCES="$APP/Contents/Resources"

[[ -d "$APP" ]] || { echo "VisualTeX.app is missing from API-only DMG" >&2; exit 1; }
codesign --verify --deep --strict "$APP"

BUNDLE_VERSION="$(/usr/libexec/PlistBuddy -c 'Print :CFBundleShortVersionString' "$APP/Contents/Info.plist")"
[[ "$BUNDLE_VERSION" == "$VERSION" ]] || {
  echo "API-only bundle version mismatch: expected $VERSION, got $BUNDLE_VERSION" >&2
  exit 1
}

for required in \
  "$RESOURCES/office/macos-offline/resources/VisualTeX.dotm" \
  "$RESOURCES/office/macos-offline/resources/VisualTeX.ppam" \
  "$RESOURCES/office/macos-offline/resources/addins.json"; do
  [[ -f "$required" ]] || { echo "Missing Office resource: $required" >&2; exit 1; }
done

for forbidden in \
  "$RESOURCES/ocr/worker.py" \
  "$RESOURCES/ocr/offline" \
  "$RESOURCES/ocr-runtime"; do
  if [[ -e "$forbidden" ]]; then
    echo "API-only DMG unexpectedly contains local OCR resource: $forbidden" >&2
    exit 1
  fi
done

APP_BYTES="$(du -sk "$APP" | awk '{print $1 * 1024}')"
DMG_SHA256="$(shasum -a 256 "$DMG_PATH" | awk '{print $1}')"
echo "Verified API-only DMG: $DMG_PATH"
echo "Bundle version: $BUNDLE_VERSION"
echo "App bytes: $APP_BYTES"
echo "Local OCR runtime/model resources: absent"
echo "Office add-ins: present"
echo "SHA-256: $DMG_SHA256"
