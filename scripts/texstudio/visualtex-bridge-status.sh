#!/bin/sh
set -eu

if [ "$#" -ne 1 ]; then
  echo "usage: visualtex-bridge-status.sh PROJECT_ROOT" >&2
  exit 64
fi

VISUALTEX_BIN=${VISUALTEX_BIN:-visualtex}
"$VISUALTEX_BIN" bridge-request "$1" initialize --params '{}' --result-only >/dev/null
exec "$VISUALTEX_BIN" bridge-status "$1"
