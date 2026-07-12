#!/bin/sh
set -eu

if [ "$#" -ne 1 ]; then
  echo "usage: visualtex-bridge-stop.sh PROJECT_ROOT" >&2
  exit 64
fi

VISUALTEX_BIN=${VISUALTEX_BIN:-visualtex}
exec "$VISUALTEX_BIN" bridge-shutdown "$1"
