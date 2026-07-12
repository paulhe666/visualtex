#!/bin/sh
set -eu

PREFIX=${VISUALTEX_TEXSTUDIO_HOME:-"$HOME/.local/share/visualtex/texstudio"}

while [ "$#" -gt 0 ]; do
  case "$1" in
    --prefix)
      [ "$#" -ge 2 ] || { echo "--prefix requires a path" >&2; exit 64; }
      PREFIX=$2
      shift 2
      ;;
    -h|--help)
      echo "usage: uninstall.sh [--prefix PATH]"
      exit 0
      ;;
    *)
      echo "unknown option: $1" >&2
      exit 64
      ;;
  esac
done

MARKER="$PREFIX/.visualtex-texstudio-adapter"
if [ ! -f "$MARKER" ]; then
  echo "refusing to remove unmarked directory: $PREFIX" >&2
  exit 73
fi

for name in \
  visualtex-bridge-ensure.sh \
  visualtex-bridge-start.sh \
  visualtex-bridge-stop.sh \
  visualtex-bridge-status.sh \
  visualtex-open.sh \
  visualtex-compile.sh \
  visualtex-forward-search.sh \
  visualtex-inverse-search.sh
do
  rm -f "$PREFIX/bin/$name"
done
rm -f "$PREFIX/README.md" "$MARKER"
rmdir "$PREFIX/bin" 2>/dev/null || true
rmdir "$PREFIX" 2>/dev/null || true

echo "VisualTeX TeXstudio adapters removed from $PREFIX"
