#!/usr/bin/env bash
set -euo pipefail

DMG_PATH="${1:-src-tauri/target/debug/bundle/dmg/VisualTeX_1.0.3_aarch64.dmg}"

if [[ "$(uname -s)" != "Darwin" ]]; then
  echo "macOS DMG verification can only run on macOS." >&2
  exit 1
fi

if [[ ! -f "$DMG_PATH" ]]; then
  echo "DMG not found: $DMG_PATH" >&2
  exit 1
fi

MOUNT_POINT="$(mktemp -d "${TMPDIR:-/tmp}/visualtex-dmg-verify.XXXXXX")"
cleanup() {
  if mount | grep -Fq "on $MOUNT_POINT "; then
    hdiutil detach "$MOUNT_POINT" >/dev/null
  fi
  rmdir "$MOUNT_POINT" 2>/dev/null || true
}
trap cleanup EXIT

hdiutil attach -nobrowse -readonly -mountpoint "$MOUNT_POINT" "$DMG_PATH" >/dev/null
APP_PATH="$MOUNT_POINT/VisualTeX.app"
CODE_RESOURCES="$APP_PATH/Contents/_CodeSignature/CodeResources"

if [[ ! -d "$APP_PATH" ]]; then
  echo "VisualTeX.app is missing from the DMG." >&2
  exit 1
fi

if [[ ! -f "$CODE_RESOURCES" ]]; then
  echo "The app bundle has no CodeResources file and will be treated as damaged after quarantine." >&2
  exit 1
fi

codesign --verify --deep --strict --verbose=2 "$APP_PATH"
IDENTIFIER="$(defaults read "$APP_PATH/Contents/Info" CFBundleIdentifier)"

if [[ "$IDENTIFIER" != "com.visualtex.studio" ]]; then
  echo "Unexpected bundle identifier: $IDENTIFIER" >&2
  exit 1
fi

echo "Verified: $DMG_PATH"
echo "Bundle identifier: $IDENTIFIER"
echo "CodeResources: present"
