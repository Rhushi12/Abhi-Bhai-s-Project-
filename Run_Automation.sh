#!/bin/bash
# ============================================
# VALUATION TEMPLATE AUTOMATION - Mac Setup
# ============================================
# Just double-click this file to run!
# First run will auto-install dependencies.
# ============================================

# Navigate to the script's directory
cd "$(dirname "$0")"

echo "============================================"
echo "  VALUATION TEMPLATE AUTOMATION"
echo "============================================"
echo ""

# Check if Python 3 is installed
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 is not installed."
    echo ""
    echo "Please install Python 3 first:"
    echo "  1. Go to: https://www.python.org/downloads/"
    echo "  2. Download and install Python 3"
    echo "  3. Run this script again"
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

echo "✓ Python 3 found"

# Check and install required packages
echo ""
echo "Checking dependencies..."

# Install pandas if not present
python3 -c "import pandas" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Installing pandas..."
    pip3 install pandas --quiet
fi

# Install openpyxl if not present
python3 -c "import openpyxl" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "Installing openpyxl..."
    pip3 install openpyxl --quiet
fi

echo "✓ All dependencies ready"
echo ""

# Run the automation script
python3 automation.py

# Keep terminal open
echo ""
read -p "Press Enter to exit..."
