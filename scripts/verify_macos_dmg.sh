#!/usr/bin/env bash
set -euo pipefail

if [[ "$(uname -s)" != "Darwin" ]]; then
  echo "macOS DMG verification can only run on macOS." >&2
  exit 1
fi

EXPECTED_VERSION="${EXPECTED_VERSION:-}"
if [[ -z "$EXPECTED_VERSION" && -f package.json ]]; then
  EXPECTED_VERSION="$(node -p "require('./package.json').version" 2>/dev/null || true)"
fi

DMG_PATH="${1:-}"
if [[ -z "$DMG_PATH" ]]; then
  DMG_PATH="$(find src-tauri/target -type f -path '*/bundle/dmg/VisualTeX_*.dmg' -print0 2>/dev/null \
    | xargs -0 ls -1t 2>/dev/null \
    | head -n 1 || true)"
fi

if [[ -z "$DMG_PATH" || ! -f "$DMG_PATH" ]]; then
  echo "VisualTeX DMG not found. Pass its path as the first argument." >&2
  exit 1
fi

# Verify the disk image container before mounting it. This catches truncated or
# corrupted uploads independently from the signature checks inside the image.
hdiutil verify "$DMG_PATH" >/dev/null

MOUNT_POINT="$(mktemp -d "${TMPDIR:-/tmp}/visualtex-dmg-verify.XXXXXX")"
cleanup() {
  if mount | grep -Fq "on $MOUNT_POINT "; then
    hdiutil detach "$MOUNT_POINT" >/dev/null || hdiutil detach -force "$MOUNT_POINT" >/dev/null
  fi
  rmdir "$MOUNT_POINT" 2>/dev/null || true
}
trap cleanup EXIT

hdiutil attach -nobrowse -readonly -mountpoint "$MOUNT_POINT" "$DMG_PATH" >/dev/null
APP_PATH="$MOUNT_POINT/VisualTeX.app"
INFO_PLIST="$APP_PATH/Contents/Info.plist"
CODE_RESOURCES="$APP_PATH/Contents/_CodeSignature/CodeResources"
EXECUTABLE="$APP_PATH/Contents/MacOS/visualtex"

if [[ ! -d "$APP_PATH" ]]; then
  echo "VisualTeX.app is missing from the DMG." >&2
  exit 1
fi
if [[ ! -f "$INFO_PLIST" ]]; then
  echo "VisualTeX.app has no Info.plist." >&2
  exit 1
fi
if [[ ! -x "$EXECUTABLE" ]]; then
  echo "VisualTeX executable is missing or not executable." >&2
  exit 1
fi
if [[ ! -f "$CODE_RESOURCES" ]]; then
  echo "The app bundle has no CodeResources file and may be treated as damaged after download." >&2
  exit 1
fi

codesign --verify --deep --strict --verbose=2 "$APP_PATH"
SIGNATURE_DETAILS="$(codesign -dvvv "$APP_PATH" 2>&1)"
IDENTIFIER="$(plutil -extract CFBundleIdentifier raw -o - "$INFO_PLIST")"
BUNDLE_VERSION="$(plutil -extract CFBundleShortVersionString raw -o - "$INFO_PLIST")"

if [[ "$IDENTIFIER" != "com.visualtex.studio" ]]; then
  echo "Unexpected bundle identifier: $IDENTIFIER" >&2
  exit 1
fi
if [[ -n "$EXPECTED_VERSION" && "$BUNDLE_VERSION" != "$EXPECTED_VERSION" ]]; then
  echo "Unexpected bundle version: $BUNDLE_VERSION (expected $EXPECTED_VERSION)" >&2
  exit 1
fi

SIGNATURE_KIND="developer-id"
if grep -q '^Signature=adhoc$' <<<"$SIGNATURE_DETAILS"; then
  SIGNATURE_KIND="ad-hoc"
fi

if [[ "${REQUIRE_NOTARIZATION:-0}" == "1" ]]; then
  if [[ "$SIGNATURE_KIND" == "ad-hoc" ]]; then
    echo "A notarized release was required, but the app only has an ad-hoc signature." >&2
    exit 1
  fi
  spctl --assess --type execute --verbose=4 "$APP_PATH"
  xcrun stapler validate "$APP_PATH"
fi

SHA256="$(shasum -a 256 "$DMG_PATH" | awk '{print $1}')"
echo "Verified DMG container: $DMG_PATH"
echo "Bundle identifier: $IDENTIFIER"
echo "Bundle version: $BUNDLE_VERSION"
echo "Signature: $SIGNATURE_KIND"
echo "CodeResources: present"
echo "SHA-256: $SHA256"

if [[ "$SIGNATURE_KIND" == "ad-hoc" ]]; then
  echo "Warning: ad-hoc signing preserves bundle integrity but cannot replace Apple Developer ID signing and notarization." >&2
fi
