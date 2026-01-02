#!/bin/bash
echo ""
echo "========================================"
echo "  Report Card Generator"
echo "  Starting server..."
echo "========================================"
echo ""

cd "$(dirname "$0")"

# Check for bundled Python or use system Python
if [ -f "python/bin/python3" ]; then
    PYTHON="python/bin/python3"
elif command -v python3 &> /dev/null; then
    PYTHON="python3"
else
    echo "ERROR: Python 3 not found!"
    echo "Please install Python 3.10+"
    exit 1
fi

# Check for venv
if [ -d "venv" ]; then
    source venv/bin/activate
fi

# Run the app
cd app
$PYTHON main.py

