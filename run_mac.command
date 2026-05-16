#!/bin/bash

echo "=========================================="
echo "Starting Tech Efficiency Analyzer"
echo "=========================================="

# Navigate to the directory where the script is located
cd "$(dirname "$0")"

# Check if Python3 is installed
if ! command -v python3 &> /dev/null; then
    echo "[ERROR] Python3 is not installed."
    echo "Please install Python3 or Xcode Command Line Tools."
    exit 1
fi

# Create virtual environment if it doesn't exist
if [ ! -d "venv" ]; then
    echo "[INFO] First time setup: Creating virtual environment..."
    python3 -m venv venv
    if [ $? -ne 0 ]; then
        echo "[ERROR] Failed to create virtual environment."
        exit 1
    fi
fi

# Activate virtual environment
source venv/bin/activate

# Install/Upgrade dependencies
echo "[INFO] Checking dependencies..."
pip install -r requirements.txt -q

# Run the dashboard
echo "[INFO] Starting the dashboard..."
streamlit run dashboard.py
