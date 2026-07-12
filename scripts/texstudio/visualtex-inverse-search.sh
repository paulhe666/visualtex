#!/bin/sh
set -eu

if [ "$#" -ne 6 ]; then
  echo "usage: visualtex-inverse-search.sh PROJECT_ROOT PDF_PATH PAGE X Y OUTPUT_JSON" >&2
  exit 64
fi

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
sh "$SCRIPT_DIR/visualtex-bridge-ensure.sh" "$1"
VISUALTEX_BIN=${VISUALTEX_BIN:-visualtex}
"$VISUALTEX_BIN" bridge-inverse-search "$1" "$2" "$3" "$4" "$5" > "$6"
