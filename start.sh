#!/bin/bash
set -e

echo "Installing dependencies..."
pip install -r requirements.txt --break-system-packages --no-cache-dir

echo "Starting application..."
python3 -m gunicorn app:app --bind 0.0.0.0:${PORT:-5000} --workers 2 --timeout 120
