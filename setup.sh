#!/bin/bash
# ─────────────────────────────────────────────────────────────
# ES Windows Form Automation - Setup Script
# ─────────────────────────────────────────────────────────────
#
# This script:
# 1. Installs Python dependencies
# 2. Launches Chrome with remote debugging enabled
#
# After running this, log into https://orders.eswindows.co
# in the Chrome window that opens, then run: python3 main.py
# ─────────────────────────────────────────────────────────────

echo "================================================"
echo "  ES Windows Form Automation - Setup"
echo "================================================"
echo

# Step 1: Install Python dependencies
echo "[1/2] Installing Python dependencies..."
pip3 install -r requirements.txt --quiet
echo "  Done."
echo

# Step 2: Launch Chrome with remote debugging
echo "[2/2] Launching Chrome with remote debugging on port 9222..."
echo "  NOTE: If Chrome is already running, close it first!"
echo

# Kill existing Chrome (optional - uncomment if needed)
# pkill -f "Google Chrome"
# sleep 2

/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome \
    --remote-debugging-port=9222 \
    --user-data-dir="/tmp/chrome-automation-profile" \
    "https://orders.eswindows.co" &

echo
echo "================================================"
echo "  Chrome is starting..."
echo "  1. Log into ES Windows in the Chrome window"
echo "  2. Then run: python3 main.py"
echo "================================================"
