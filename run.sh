#!/bin/bash
set -e

echo "================================"
echo "  Montblanc Dashboard - Setup"
echo "================================"

REPO_DIR="$HOME/For-Claude"
APP_DIR="$REPO_DIR/montblanc-dashboard"

# Clone or update repo
if [ -d "$REPO_DIR" ]; then
  echo "Updating..."
  git -C "$REPO_DIR" pull --quiet
else
  echo "Downloading app..."
  git clone --quiet https://github.com/Btriaire/For-Claude.git "$REPO_DIR"
fi

cd "$APP_DIR"

# Install dependencies
echo "Installing dependencies..."
pip3 install -r requirements.txt --quiet

# Get local IP for iPad access
IP=$(ipconfig getifaddr en0 2>/dev/null || ipconfig getifaddr en1 2>/dev/null || echo "")

echo ""
echo "================================"
echo "  App running!"
echo ""
echo "  Mac :  http://localhost:5001"
[ -n "$IP" ] && echo "  iPad : http://$IP:5001"
echo "================================"
echo ""
echo "Press Ctrl+C to stop."
echo ""

python3 app.py
