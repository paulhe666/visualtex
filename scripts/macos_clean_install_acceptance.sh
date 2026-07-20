#!/usr/bin/env bash
set -euo pipefail

if [[ "$(uname -s)" != "Darwin" ]]; then
  echo "VisualTeX clean-install acceptance can only run on macOS." >&2
  exit 1
fi

PROJECT_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
DMG_PATH="${1:-$PROJECT_ROOT/src-tauri/target/release/bundle/dmg/VisualTeX_1.1.0_aarch64.dmg}"
APP_PATH="/Applications/VisualTeX.app"
EXECUTABLE="$APP_PATH/Contents/MacOS/visualtex"
OFFICE_RESOURCES="$APP_PATH/Contents/Resources/office/macos-offline/resources"
GROUP_ROOT="$HOME/Library/Group Containers/UBF8T346G9.Office/VisualTeX"
WORD_STARTUP_ROOT="$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word"
WORD_SCRIPTS_ROOT="$HOME/Library/Application Scripts/com.microsoft.Word"
POWERPOINT_SCRIPTS_ROOT="$HOME/Library/Application Scripts/com.microsoft.Powerpoint"
WORD_RUNTIME="$WORD_SCRIPTS_ROOT/VisualTeXRuntime"
POWERPOINT_RUNTIME="$POWERPOINT_SCRIPTS_ROOT/VisualTeXRuntime"
POWERPOINT_ADDIN="$GROUP_ROOT/OfficeAddins/VisualTeX.ppam"
WORD_PLACEHOLDER_COMPAT="$WORD_RUNTIME/OfficeAddins/resources/placeholder.png"
WORD_PLACEHOLDER_CANONICAL="$WORD_SCRIPTS_ROOT/VisualTeXPlaceholder.png"
LAUNCH_AGENT_LABEL="com.visualtex.studio.office"
LAUNCH_AGENT="$HOME/Library/LaunchAgents/$LAUNCH_AGENT_LABEL.plist"
APP_SUPPORT="$HOME/Library/Application Support/com.visualtex.studio"
MOUNT_POINT="$(mktemp -d "${TMPDIR:-/tmp}/visualtex-clean-install.XXXXXX")"

cleanup_mount() {
  # hdiutil canonicalizes /var to /private/var, so always attempt detach by
  # mount point instead of relying on a literal mount-output comparison.
  hdiutil detach "$MOUNT_POINT" >/dev/null 2>&1 \
    || hdiutil detach -force "$MOUNT_POINT" >/dev/null 2>&1 \
    || true
  rm -rf "$MOUNT_POINT"
}
trap cleanup_mount EXIT

require_file() {
  local path="$1"
  local label="$2"
  if [[ ! -f "$path" || ! -s "$path" ]]; then
    echo "$label is missing or empty: $path" >&2
    exit 1
  fi
}

require_absent() {
  local path="$1"
  local label="$2"
  if [[ -e "$path" ]]; then
    echo "$label still exists after cleanup: $path" >&2
    exit 1
  fi
}

require_hash_match() {
  local source="$1"
  local installed="$2"
  local label="$3"
  require_file "$source" "$label package source"
  require_file "$installed" "$label installed copy"
  local source_hash installed_hash
  source_hash="$(shasum -a 256 "$source" | awk '{print $1}')"
  installed_hash="$(shasum -a 256 "$installed" | awk '{print $1}')"
  if [[ "$source_hash" != "$installed_hash" ]]; then
    echo "$label hash mismatch: package=$source_hash installed=$installed_hash" >&2
    exit 1
  fi
  echo "$label SHA-256: $installed_hash"
}

if [[ ! -f "$DMG_PATH" ]]; then
  echo "VisualTeX DMG not found: $DMG_PATH" >&2
  exit 1
fi
DMG_PATH="$(cd "$(dirname "$DMG_PATH")" && pwd)/$(basename "$DMG_PATH")"

for process_name in "Microsoft Word" "Microsoft PowerPoint"; do
  if pgrep -x "$process_name" >/dev/null 2>&1; then
    echo "$process_name is still running. Fully quit it with Command-Q before clean-install acceptance." >&2
    exit 1
  fi
done

BACKUP_ROOT="$GROUP_ROOT/Scratch/CleanInstallBackup/$(date +%Y%m%d-%H%M%S)"
mkdir -p "$BACKUP_ROOT"
for artifact in \
  "$PROJECT_ROOT/office/macos-offline/resources/VisualTeX.dotm" \
  "$PROJECT_ROOT/office/macos-offline/resources/VisualTeX.ppam"; do
  require_file "$artifact" "Reviewed Office artifact"
  cp -p "$artifact" "$BACKUP_ROOT/$(basename "$artifact")"
done

echo "Reviewed Office artifacts backed up to: $BACKUP_ROOT"

pkill -x visualtex >/dev/null 2>&1 || true
for _ in {1..20}; do
  if ! pgrep -x visualtex >/dev/null 2>&1; then
    break
  fi
  sleep 0.1
done
if pgrep -x visualtex >/dev/null 2>&1; then
  echo "VisualTeX processes did not stop before cleanup." >&2
  exit 1
fi

USER_DOMAIN="gui/$(id -u)"
launchctl bootout "$USER_DOMAIN/$LAUNCH_AGENT_LABEL" >/dev/null 2>&1 || true
launchctl disable "$USER_DOMAIN/$LAUNCH_AGENT_LABEL" >/dev/null 2>&1 || true
rm -f "$LAUNCH_AGENT"
rm -f \
  "$APP_SUPPORT/office-background.enabled" \
  "$APP_SUPPORT/office-background.enabled.disabled-for-dev" \
  "$APP_SUPPORT/dock-icon-v2.refreshed"

rm -rf "$APP_PATH" "$HOME/Applications/VisualTeX.app"

# Preserve the large offline OCR runtime, but remove retired Office.js state,
# old Office sessions, certificates, recovery data, and packaged add-in copies.
rm -rf \
  "$APP_SUPPORT/office" \
  "$APP_SUPPORT/OfficeAddins" \
  "$HOME/Library/Application Support/VisualTeX" \
  "$HOME/Library/Caches/com.visualtex.studio" \
  "$HOME/Library/Caches/visualtex" \
  "$HOME/Library/WebKit/com.visualtex.studio" \
  "$HOME/Library/WebKit/visualtex" \
  "$HOME/Library/HTTPStorages/com.visualtex.studio" \
  "$HOME/Library/Saved Application State/com.visualtex.studio.savedState" \
  "$HOME/Library/Logs/VisualTeX"
rm -f "$HOME/Library/Preferences/com.visualtex.studio.plist"

rm -rf "$WORD_RUNTIME" "$POWERPOINT_RUNTIME"
rm -f \
  "$WORD_SCRIPTS_ROOT/VisualTeXWord.scpt" \
  "$WORD_PLACEHOLDER_CANONICAL" \
  "$POWERPOINT_SCRIPTS_ROOT/VisualTeXPowerPoint.scpt"

if [[ -d "$WORD_STARTUP_ROOT" ]]; then
  find "$WORD_STARTUP_ROOT" -maxdepth 1 -type f -iname 'VisualTeX.dotm*' -delete
fi
rm -rf \
  "$GROUP_ROOT/OfficeAddins" \
  "$GROUP_ROOT/OfficePluginStatus" \
  "$GROUP_ROOT/OfficeSessions" \
  "$GROUP_ROOT/NativeDocuments" \
  "$GROUP_ROOT/Tests"

for host_container in com.microsoft.Word com.microsoft.Powerpoint; do
  WEF_ROOT="$HOME/Library/Containers/$host_container/Data/Documents/wef"
  if [[ -d "$WEF_ROOT" ]]; then
    find "$WEF_ROOT" -maxdepth 6 -type f \( \
      -iname '*d6fcb260-4c37-4f73-a173-cf24674f81f2*' -o \
      -iname '*a6d13cf2-54e8-4dfa-a20c-15de864ab3c5*' -o \
      -iname 'visualtex.word.xml*' -o \
      -iname 'visualtex.powerpoint.xml*' \
    \) -delete
  fi
done

require_absent "$APP_PATH" "Old VisualTeX application"
require_absent "$LAUNCH_AGENT" "Old VisualTeX LaunchAgent"
require_absent "$WORD_RUNTIME" "Old Word runtime"
require_absent "$POWERPOINT_RUNTIME" "Old PowerPoint runtime"
require_absent "$POWERPOINT_ADDIN" "Old fixed PowerPoint add-in"

hdiutil attach -nobrowse -readonly -mountpoint "$MOUNT_POINT" "$DMG_PATH" >/dev/null
require_file "$MOUNT_POINT/VisualTeX.app/Contents/MacOS/visualtex" "VisualTeX app in DMG"
ditto "$MOUNT_POINT/VisualTeX.app" "$APP_PATH"
hdiutil detach "$MOUNT_POINT" >/dev/null
rmdir "$MOUNT_POINT" 2>/dev/null || true

require_file "$EXECUTABLE" "Installed VisualTeX executable"
codesign --verify --deep --strict --verbose=2 "$APP_PATH"

LSREGISTER="/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister"
if [[ -x "$LSREGISTER" ]]; then
  "$LSREGISTER" -f "$APP_PATH" >/dev/null 2>&1 || true
fi

if pgrep -x visualtex >/dev/null 2>&1; then
  echo "VisualTeX unexpectedly started before native add-in installation." >&2
  exit 1
fi

"$EXECUTABLE" --install-macos-office-addins

WORD_ADDIN="$WORD_STARTUP_ROOT/VisualTeX.dotm"
WORD_SCRIPT="$WORD_SCRIPTS_ROOT/VisualTeXWord.scpt"
POWERPOINT_SCRIPT="$POWERPOINT_SCRIPTS_ROOT/VisualTeXPowerPoint.scpt"

require_hash_match "$OFFICE_RESOURCES/VisualTeX.dotm" "$WORD_ADDIN" "Word DOTM"
require_hash_match "$OFFICE_RESOURCES/VisualTeX.ppam" "$POWERPOINT_ADDIN" "PowerPoint PPAM"
require_file "$WORD_SCRIPT" "Word AppleScriptTask"
require_file "$POWERPOINT_SCRIPT" "PowerPoint AppleScriptTask"
require_file "$WORD_PLACEHOLDER_COMPAT" "Compiled-DOTM compatibility placeholder"
require_file "$WORD_PLACEHOLDER_CANONICAL" "Canonical Word placeholder"

if launchctl print "$USER_DOMAIN/$LAUNCH_AGENT_LABEL" >/dev/null 2>&1; then
  echo "A VisualTeX LaunchAgent was unexpectedly recreated during native add-in installation." >&2
  exit 1
fi
require_absent "$LAUNCH_AGENT" "VisualTeX LaunchAgent plist"

if find "$WORD_RUNTIME/OfficeSessions" "$POWERPOINT_RUNTIME/OfficeSessions" -mindepth 1 -print -quit 2>/dev/null | grep -q .; then
  echo "A stale Office Session survived the clean installation." >&2
  exit 1
fi

open "$APP_PATH"
for _ in {1..50}; do
  if pgrep -x visualtex >/dev/null 2>&1; then
    break
  fi
  sleep 0.1
done

VISUALTEX_COMMAND="$(ps -p "$(pgrep -x visualtex | head -n 1)" -o command= 2>/dev/null || true)"
if [[ -z "$VISUALTEX_COMMAND" ]]; then
  echo "VisualTeX did not start after clean installation." >&2
  exit 1
fi
if [[ "$VISUALTEX_COMMAND" == *"--office-background"* ]]; then
  echo "Foreground VisualTeX incorrectly started in Office background mode: $VISUALTEX_COMMAND" >&2
  exit 1
fi

DOCK_SIZE=""
for _ in {1..50}; do
  DOCK_SIZE="$(osascript <<'APPLESCRIPT' 2>/dev/null || true
tell application "System Events"
  tell process "Dock"
    repeat with dockItem in UI elements of list 1
      try
        if (name of dockItem as text) is "VisualTeX" then
          set itemSize to size of dockItem
          return (item 1 of itemSize) & "x" & (item 2 of itemSize)
        end if
      end try
    end repeat
  end tell
end tell
APPLESCRIPT
)"
  if [[ -n "$DOCK_SIZE" ]]; then
    break
  fi
  sleep 0.1
done

DOCK_WIDTH="${DOCK_SIZE%x*}"
DOCK_HEIGHT="${DOCK_SIZE#*x}"
if [[ -z "$DOCK_SIZE" || ! "$DOCK_WIDTH" =~ ^[0-9]+$ || ! "$DOCK_HEIGHT" =~ ^[0-9]+$ \
      || "$DOCK_WIDTH" -lt 30 || "$DOCK_HEIGHT" -lt 40 ]]; then
  echo "VisualTeX Dock tile is missing or too small after migration refresh: ${DOCK_SIZE:-missing}" >&2
  exit 1
fi

require_file "$APP_SUPPORT/dock-icon-v2.refreshed" "Dock icon migration marker"

echo "VisualTeX clean-install acceptance: PASS"
echo "Installed app: $APP_PATH"
echo "Foreground command: $VISUALTEX_COMMAND"
echo "Dock tile size: $DOCK_SIZE"
echo "DMG SHA-256: $(shasum -a 256 "$DMG_PATH" | awk '{print $1}')"
