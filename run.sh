#!/bin/bash
set -e

echo "================================"
echo "  Montblanc Dashboard - Setup"
echo "================================"

# Clone or update repo
if [ -d "$HOME/montblanc-dashboard" ]; then
  echo "Updating existing installation..."
  cd "$HOME/montblanc-dashboard"
  git pull --quiet
else
  echo "Downloading app..."
  git clone --quiet https://github.com/Btriaire/For-Claude.git "$HOME/montblanc-dashboard"
  cd "$HOME/montblanc-dashboard"
fi

cd montblanc-dashboard

# Install dependencies
echo "Installing dependencies..."
pip3 install -r requirements.txt --quiet

# Get local IP for iPad access
IP=$(ipconfig getifaddr en0 2>/dev/null || ipconfig getifaddr en1 2>/dev/null || echo "unknown")

echo ""
echo "================================"
echo "  App running!"
echo ""
echo "  Mac :  http://localhost:5000"
if [ "$IP" != "unknown" ]; then
echo "  iPad : http://$IP:5000"
fi
echo "================================"
echo ""
echo "Press Ctrl+C to stop."
echo ""

python3 app.py
