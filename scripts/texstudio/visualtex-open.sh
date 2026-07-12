#!/bin/sh
set -eu

if [ "$#" -ne 1 ]; then
  echo "usage: visualtex-open.sh PROJECT_ROOT" >&2
  exit 64
fi

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
sh "$SCRIPT_DIR/visualtex-bridge-ensure.sh" "$1"
VISUALTEX_BIN=${VISUALTEX_BIN:-visualtex}
exec "$VISUALTEX_BIN" open "$1"
