#!/bin/bash
# Propellic Claude Skills — one-time setup script
# Run this once: bash setup.sh

set -e

PROPELLIC_DIR="$HOME/.propellic"
VENV="$PROPELLIC_DIR/venv"
CREDS="$PROPELLIC_DIR/credentials.json"
TOKEN="$PROPELLIC_DIR/token.json"

echo ""
echo "╔══════════════════════════════════════════════╗"
echo "║   Propellic Claude Skills — Setup            ║"
echo "╚══════════════════════════════════════════════╝"
echo ""

# 1. Create ~/.propellic if needed
mkdir -p "$PROPELLIC_DIR"
echo "✓ Created $PROPELLIC_DIR"

# 2. Check Python 3
if ! command -v python3 &>/dev/null; then
  echo "✗ Python 3 not found. Install it from https://python.org and re-run."
  exit 1
fi
echo "✓ Python 3 found: $(python3 --version)"

# 3. Create venv
if [ ! -d "$VENV" ]; then
  python3 -m venv "$VENV"
  echo "✓ Created venv at $VENV"
else
  echo "✓ Venv already exists at $VENV"
fi

# 4. Install packages
"$VENV/bin/pip" install --quiet --upgrade pip
"$VENV/bin/pip" install --quiet \
  python-docx \
  python-pptx \
  openpyxl \
  google-api-python-client \
  google-auth-httplib2 \
  google-auth-oauthlib
echo "✓ Installed all required Python packages"

# 5. Check for credentials.json
if [ ! -f "$CREDS" ]; then
  echo ""
  echo "──────────────────────────────────────────────"
  echo "  Google credentials not found."
  echo ""
  echo "  Ask Javier or Eric for the credentials.json"
  echo "  file and place it at:"
  echo ""
  echo "    $CREDS"
  echo ""
  echo "  Then re-run this script to complete setup."
  echo "──────────────────────────────────────────────"
  exit 1
fi
echo "✓ Found Google credentials"

# 6. Run auth flow if token doesn't exist
if [ ! -f "$TOKEN" ]; then
  echo ""
  echo "Opening browser for Google authorization..."
  "$VENV/bin/python3" "$(dirname "$0")/propellic-auth.py"
else
  echo "✓ Google token already exists — checking validity..."
  "$VENV/bin/python3" - <<'EOF'
import os
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request

token_path = os.path.expanduser('~/.propellic/token.json')
creds = Credentials.from_authorized_user_file(token_path)
if creds.expired and creds.refresh_token:
    creds.refresh(Request())
    with open(token_path, 'w') as f:
        f.write(creds.to_json())
    print("✓ Token refreshed")
else:
    print("✓ Token is valid")
EOF
fi

echo ""
echo "╔══════════════════════════════════════════════╗"
echo "║   Setup complete! You're ready to use        ║"
echo "║   /propellic-doc, /propellic-sheets,         ║"
echo "║   and /propellic-slides in Claude Code.      ║"
echo "╚══════════════════════════════════════════════╝"
echo ""
