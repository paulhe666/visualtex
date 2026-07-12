#!/bin/sh
set -eu

PREFIX=${VISUALTEX_TEXSTUDIO_HOME:-"$HOME/.local/share/visualtex/texstudio"}
FORCE=0

while [ "$#" -gt 0 ]; do
  case "$1" in
    --prefix)
      [ "$#" -ge 2 ] || { echo "--prefix requires a path" >&2; exit 64; }
      PREFIX=$2
      shift 2
      ;;
    --force)
      FORCE=1
      shift
      ;;
    -h|--help)
      echo "usage: install.sh [--prefix PATH] [--force]"
      exit 0
      ;;
    *)
      echo "unknown option: $1" >&2
      exit 64
      ;;
  esac
done

SOURCE_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
BIN_DIR="$PREFIX/bin"
mkdir -p "$BIN_DIR"

FILES="
visualtex-bridge-ensure.sh
visualtex-bridge-start.sh
visualtex-bridge-stop.sh
visualtex-bridge-status.sh
visualtex-open.sh
visualtex-compile.sh
visualtex-forward-search.sh
visualtex-inverse-search.sh
"

for name in $FILES; do
  source="$SOURCE_DIR/$name"
  destination="$BIN_DIR/$name"
  if [ -e "$destination" ] && [ "$FORCE" -ne 1 ]; then
    if cmp -s "$source" "$destination"; then
      chmod 755 "$destination"
      continue
    fi
    echo "refusing to overwrite existing adapter: $destination" >&2
    echo "rerun with --force only if this is an older VisualTeX adapter" >&2
    exit 73
  fi
  temporary="$destination.tmp.$$"
  cp "$source" "$temporary"
  chmod 755 "$temporary"
  mv -f "$temporary" "$destination"
done

printf '%s\n' "VisualTeX TeXstudio adapter v1" > "$PREFIX/.visualtex-texstudio-adapter"
cp "$SOURCE_DIR/README.md" "$PREFIX/README.md"

cat <<EOF
VisualTeX TeXstudio adapters installed in:
  $BIN_DIR

This installer did not modify TeXstudio settings.
Add user commands in TeXstudio and point them at the scripts above.
See:
  $PREFIX/README.md
EOF
