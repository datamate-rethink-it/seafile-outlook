#!/bin/bash
# Generate self-signed certificate for the HTTPS dev server.
# The certificate is valid for localhost only.

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

openssl req -x509 -newkey rsa:2048 -nodes \
  -keyout "$SCRIPT_DIR/key.pem" \
  -out "$SCRIPT_DIR/cert.pem" \
  -days 365 \
  -subj "/CN=localhost" \
  -addext "subjectAltName=DNS:localhost,IP:127.0.0.1"

echo ""
echo "Certificates generated:"
echo "  $SCRIPT_DIR/cert.pem"
echo "  $SCRIPT_DIR/key.pem"
echo ""
echo "Next: node dev/server.js"
