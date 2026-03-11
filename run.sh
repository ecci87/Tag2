#!/usr/bin/env bash
# Run the Image Captioning Tool server
set -e
cd "$(dirname "$0")"

if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
fi

source venv/bin/activate
pip install -q -r requirements.txt
echo ""
echo "=== Image Captioning Tool ==="
echo "Open http://localhost:8899 in your browser"
echo ""
python server.py "$@"
