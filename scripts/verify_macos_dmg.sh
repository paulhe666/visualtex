#!/usr/bin/env bash
set -euo pipefail

if [[ "$(uname -s)" != "Darwin" ]]; then
  echo "macOS DMG verification can only run on macOS." >&2
  exit 1
fi

PROJECT_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
EXPECTED_VERSION="${EXPECTED_VERSION:-}"
if [[ -z "$EXPECTED_VERSION" && -f "$PROJECT_ROOT/package.json" ]]; then
  EXPECTED_VERSION="$(node -p "require('$PROJECT_ROOT/package.json').version" 2>/dev/null || true)"
fi

DMG_PATH="${1:-}"
if [[ -z "$DMG_PATH" ]]; then
  DMG_PATH="$(find "$PROJECT_ROOT/src-tauri/target" -type f -path '*/bundle/dmg/VisualTeX_*.dmg' -print0 2>/dev/null \
    | xargs -0 ls -1t 2>/dev/null \
    | head -n 1 || true)"
fi

if [[ -z "$DMG_PATH" || ! -f "$DMG_PATH" ]]; then
  echo "VisualTeX DMG not found. Pass its path as the first argument." >&2
  exit 1
fi
DMG_PATH="$(cd "$(dirname "$DMG_PATH")" && pwd)/$(basename "$DMG_PATH")"

require_file() {
  local path="$1"
  local label="$2"
  if [[ ! -f "$path" || ! -s "$path" ]]; then
    echo "$label is missing or empty: $path" >&2
    exit 1
  fi
}

require_directory() {
  local path="$1"
  local label="$2"
  if [[ ! -d "$path" ]]; then
    echo "$label is missing: $path" >&2
    exit 1
  fi
}

hdiutil verify "$DMG_PATH" >/dev/null

MOUNT_POINT="$(mktemp -d "${TMPDIR:-/tmp}/visualtex-dmg-verify.XXXXXX")"
ICON_CHECK_DIR="$(mktemp -d "${TMPDIR:-/tmp}/visualtex-icon-verify.XXXXXX")"
cleanup() {
  # hdiutil canonicalizes /var to /private/var, so checking the literal mount
  # string can miss a live image. Always attempt detach by mount point.
  hdiutil detach "$MOUNT_POINT" >/dev/null 2>&1 \
    || hdiutil detach -force "$MOUNT_POINT" >/dev/null 2>&1 \
    || true
  rm -rf "$MOUNT_POINT" "$ICON_CHECK_DIR"
}
trap cleanup EXIT

hdiutil attach -nobrowse -readonly -mountpoint "$MOUNT_POINT" "$DMG_PATH" >/dev/null
APP_PATH="$MOUNT_POINT/VisualTeX.app"
INFO_PLIST="$APP_PATH/Contents/Info.plist"
CODE_RESOURCES="$APP_PATH/Contents/_CodeSignature/CodeResources"
EXECUTABLE="$APP_PATH/Contents/MacOS/visualtex"
RESOURCES="$APP_PATH/Contents/Resources"
APP_ICON="$RESOURCES/icon.icns"
OFFICE_ROOT="$RESOURCES/office/macos-offline"
OCR_ROOT="$RESOURCES/ocr/offline/macos-arm64"

require_directory "$APP_PATH" "VisualTeX.app"
require_file "$INFO_PLIST" "Info.plist"
require_file "$CODE_RESOURCES" "CodeResources"
if [[ ! -x "$EXECUTABLE" ]]; then
  echo "VisualTeX executable is missing or not executable: $EXECUTABLE" >&2
  exit 1
fi
require_directory "$RESOURCES" "Application Resources"
require_file "$APP_ICON" "macOS application icon"
ICON_CHECK_PNG="$ICON_CHECK_DIR/icon-32.png"
sips -s format png -z 32 32 "$APP_ICON" --out "$ICON_CHECK_PNG" >/dev/null
python3 "$PROJECT_ROOT/scripts/verify_macos_app_icon.py" "$ICON_CHECK_PNG" \
  --expected-subject-rgb 31,99,142 \
  --subject-rgb-tolerance 24 \
  --minimum-subject-pixels 8 \
  --minimum-white-ratio 0.40

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

ARCHITECTURES="$(lipo -archs "$EXECUTABLE")"
if ! grep -Eq '(^|[[:space:]])arm64($|[[:space:]])' <<<"$ARCHITECTURES"; then
  echo "VisualTeX executable does not contain arm64: $ARCHITECTURES" >&2
  exit 1
fi

# macOS Office resources must contain only the native DOTM/PPAM integration.
for relative in \
  PROTOCOL.md \
  POWERPOINT_INSTALL.md \
  resources/VisualTeX.dotm \
  resources/VisualTeX.ppam \
  resources/addins.json \
  word/VisualTeXWord.scpt \
  word/VTWordAdapter.bas \
  word/customUI14.xml \
  powerpoint/VisualTeXPowerPoint.scpt \
  powerpoint/VTPowerPointAdapter.bas \
  powerpoint/customUI14.xml; do
  require_file "$OFFICE_ROOT/$relative" "Native Office resource $relative"
done
xmllint --noout \
  "$OFFICE_ROOT/word/customUI14.xml" \
  "$OFFICE_ROOT/powerpoint/customUI14.xml"

if [[ -e "$RESOURCES/office/bridge" || -e "$RESOURCES/office/dialog" || -e "$RESOURCES/office/vendor" || -e "$RESOURCES/office/manifests" ]]; then
  echo "The macOS bundle still contains the obsolete Office.js web integration." >&2
  exit 1
fi
if grep -R -I -n -E 'office\.js|SourceLocation|trusted catalog|localhost:43127' "$OFFICE_ROOT" >/tmp/visualtex-office-bundle-scan.txt 2>/dev/null; then
  cat /tmp/visualtex-office-bundle-scan.txt >&2
  rm -f /tmp/visualtex-office-bundle-scan.txt
  echo "Native Office resources contain an obsolete web integration marker." >&2
  exit 1
fi
rm -f /tmp/visualtex-office-bundle-scan.txt

# Offline OCR resources and notices are mandatory in the complete macOS build.
require_file "$RESOURCES/ocr/worker.py" "OCR worker"
for relative in \
  manifest.json \
  runtime-python310-macos-arm64.tar.gz \
  model-PP-FormulaNet_plus-M.tar.gz \
  licenses/CPYTHON-LICENSE.txt \
  licenses/PYTHON-BUILD-STANDALONE-LICENSE.txt \
  licenses/THIRD_PARTY_NOTICES.json; do
  require_file "$OCR_ROOT/$relative" "Offline OCR resource $relative"
done
python3 "$PROJECT_ROOT/scripts/verify_offline_ocr_bundle.py" --bundle "$OCR_ROOT"
if [[ "${FULL_OCR_VERIFY:-0}" == "1" ]]; then
  python3 "$PROJECT_ROOT/scripts/verify_offline_ocr_bundle.py" --bundle "$OCR_ROOT" --full
fi

# User-generated trust material and session state must never be shipped.
PRIVATE_MATCHES="$(find "$RESOURCES" \
  \( -type f \( \
    -name 'localhost-key.pem' -o \
    -name 'localhost-cert.pem' -o \
    -name 'certificate.json' -o \
    -name 'install.json' -o \
    -name 'session.json' \
  \) -o -type d \( \
    -name sessions -o \
    -name recovery -o \
    -name formulas \
  \) \) -print 2>/dev/null || true)"
if [[ -n "$PRIVATE_MATCHES" ]]; then
  echo "The app bundle contains private runtime state:" >&2
  echo "$PRIVATE_MATCHES" >&2
  exit 1
fi

# Debug fallbacks and local build paths must be compiled out of the release.
if grep -a -Fq '/Users/lpj/devspace/workspaces/visualtex' "$EXECUTABLE"; then
  echo "The release executable contains the developer workspace path." >&2
  exit 1
fi
LOCAL_PATH_MATCHES="$(grep -R -I -l -F '/Users/lpj/devspace/workspaces/visualtex' \
  "$OFFICE_ROOT" "$RESOURCES/ocr/worker.py" 2>/dev/null || true)"
if [[ -n "$LOCAL_PATH_MATCHES" ]]; then
  echo "Bundled text resources contain the developer workspace path:" >&2
  echo "$LOCAL_PATH_MATCHES" >&2
  exit 1
fi

# Confirm the background lifecycle made it into the built executable.
if ! grep -a -Fq -- '--office-background' "$EXECUTABLE"; then
  echo "The release executable has no Office background entry point." >&2
  exit 1
fi
if ! grep -a -Fq 'com.visualtex.studio.office' "$EXECUTABLE"; then
  echo "The release executable has no VisualTeX LaunchAgent label." >&2
  exit 1
fi
if ! grep -a -Fq -- '--install-macos-office-addins' "$EXECUTABLE"; then
  echo "The release executable has no clean-install native Office maintenance entry point." >&2
  exit 1
fi
if ! grep -a -Fq 'dock-icon-v2.refreshed' "$EXECUTABLE"; then
  echo "The release executable has no one-time Dock icon migration refresh." >&2
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
RESOURCE_BYTES="$(du -sk "$RESOURCES" | awk '{print $1 * 1024}')"
OCR_BYTES="$(du -sk "$OCR_ROOT" | awk '{print $1 * 1024}')"
OFFICE_FILES="$(find "$OFFICE_ROOT" -type f | wc -l | tr -d ' ')"

echo "Verified DMG container: $DMG_PATH"
echo "Bundle identifier: $IDENTIFIER"
echo "Bundle version: $BUNDLE_VERSION"
echo "Executable architectures: $ARCHITECTURES"
echo "Signature: $SIGNATURE_KIND"
echo "CodeResources: present"
echo "Dock icon visibility: verified at 32x32"
echo "Office files: $OFFICE_FILES"
echo "Resources bytes: $RESOURCE_BYTES"
echo "Offline OCR bytes: $OCR_BYTES"
echo "SHA-256: $SHA256"

if [[ "$SIGNATURE_KIND" == "ad-hoc" ]]; then
  echo "Warning: ad-hoc signing preserves bundle integrity but cannot replace Apple Developer ID signing and notarization." >&2
fi
