#!/bin/bash
# Creates VisualTeX's own self-signed macOS code-signing identity, once.
#
# macOS remembers a user's "access data from other apps" answer for the app's
# designated requirement. An ad-hoc signature's requirement is the exact build,
# so every update was denied without a prompt. Signed with this identity, the
# requirement is "com.visualtex.studio signed by this certificate", which every
# release signed with it meets. No Apple developer account is involved, and no
# Mac has to trust the certificate.
#
# The identity is a PKCS#12 file plus its password in
# ${VISUALTEX_MACOS_SIGNING_DIR:-~/.config/visualtex/macos-signing}. It never
# goes into git. Keep a backup: a release signed with another identity is
# denied once more on every user's Mac (VisualTeX then resets the record and
# macOS asks again). tauri_build.mjs signs with it from a temporary keychain.
set -euo pipefail

NAME="VisualTeX Code Signing"
DIR="${VISUALTEX_MACOS_SIGNING_DIR:-$HOME/.config/visualtex/macos-signing}"
P12="$DIR/identity.p12"
PASSWORD="$DIR/identity.password"
OPENSSL=/usr/bin/openssl # LibreSSL: its PKCS#12 encryption is what `security import` reads

if [ -e "$P12" ] || [ -e "$PASSWORD" ]; then
  echo "The VisualTeX signing identity already exists in $DIR." >&2
  echo "Never replace it: every release must be signed with the same identity." >&2
  exit 1
fi

umask 077
mkdir -p "$DIR"
chmod 700 "$DIR"
WORK=$(mktemp -d)
trap 'rm -rf "$WORK"' EXIT

cat >"$WORK/certificate.cnf" <<EOF
[req]
distinguished_name = subject
prompt = no
x509_extensions = codesigning

[subject]
CN = $NAME
O = VisualTeX

[codesigning]
basicConstraints = critical, CA:false
keyUsage = critical, digitalSignature
extendedKeyUsage = critical, codeSigning
subjectKeyIdentifier = hash
EOF

"$OPENSSL" req -new -x509 -newkey rsa:3072 -nodes -sha256 -days 7300 \
  -config "$WORK/certificate.cnf" \
  -keyout "$WORK/key.pem" -out "$WORK/certificate.pem" 2>/dev/null
"$OPENSSL" rand -hex 24 >"$PASSWORD"
"$OPENSSL" pkcs12 -export -name "$NAME" \
  -inkey "$WORK/key.pem" -in "$WORK/certificate.pem" \
  -passout "file:$PASSWORD" -out "$P12"
chmod 600 "$P12" "$PASSWORD"

echo "Created the VisualTeX signing identity in $DIR"
"$OPENSSL" x509 -in "$WORK/certificate.pem" -noout -subject -enddate -fingerprint -sha1
echo "Back up identity.p12 and identity.password together, somewhere private."
