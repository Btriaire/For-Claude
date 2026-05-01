#!/bin/bash
set -e

echo "================================"
echo "  Montblanc Dashboard - Setup"
echo "================================"

# Go to script directory
cd "$(dirname "$0")/montblanc-dashboard"

# Install dependencies
echo ""
echo "Installing dependencies..."
pip3 install -r requirements.txt --quiet

# Get local IP for iPad access
IP=$(ipconfig getifaddr en0 2>/dev/null || ipconfig getifaddr en1 2>/dev/null || echo "unknown")

echo ""
echo "================================"
echo "  App running!"
echo ""
echo "  Mac:   http://localhost:5000"
if [ "$IP" != "unknown" ]; then
echo "  iPad:  http://$IP:5000"
fi
echo "================================"
echo ""
echo "Press Ctrl+C to stop."
echo ""

python3 app.py
