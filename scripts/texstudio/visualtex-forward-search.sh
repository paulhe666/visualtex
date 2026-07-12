#!/bin/sh
set -eu

if [ "$#" -ne 6 ]; then
  echo "usage: visualtex-forward-search.sh PROJECT_ROOT SOURCE_FILE LINE COLUMN PDF_PATH OUTPUT_JSON" >&2
  exit 64
fi

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
sh "$SCRIPT_DIR/visualtex-bridge-ensure.sh" "$1"
VISUALTEX_BIN=${VISUALTEX_BIN:-visualtex}
"$VISUALTEX_BIN" bridge-forward-search "$1" "$2" "$3" "$4" "$5" > "$6"
