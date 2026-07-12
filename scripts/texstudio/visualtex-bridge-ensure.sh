#!/bin/sh
set -eu

if [ "$#" -ne 1 ]; then
  echo "usage: visualtex-bridge-ensure.sh PROJECT_ROOT" >&2
  exit 64
fi

PROJECT_ROOT=$1
VISUALTEX_BIN=${VISUALTEX_BIN:-visualtex}
SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)

if "$VISUALTEX_BIN" bridge-request "$PROJECT_ROOT" initialize --params '{}' >/dev/null 2>&1; then
  exit 0
fi

BRIDGE_DIR="$PROJECT_ROOT/.visualtex/bridge"
LOG_FILE="$BRIDGE_DIR/texstudio-bridge.log"
mkdir -p "$BRIDGE_DIR"
nohup "$VISUALTEX_BIN" bridge-serve "$PROJECT_ROOT" >"$LOG_FILE" 2>&1 &

attempt=0
while [ "$attempt" -lt 50 ]; do
  if "$VISUALTEX_BIN" bridge-request "$PROJECT_ROOT" initialize --params '{}' >/dev/null 2>&1; then
    exit 0
  fi
  attempt=$((attempt + 1))
  sleep 0.1
done

echo "VisualTeX bridge did not become ready. See: $LOG_FILE" >&2
exit 1
