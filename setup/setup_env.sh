#!/usr/bin/env bash

# Exit immediately if any command fails
set -e

# Create virtual environment
python3 -m venv venv

# Activate virtual environment
source venv/bin/activate

# Upgrade pip
pip install --upgrade pip

# Install dependencies
pip install -r requirements.txt

echo "✅ Virtual environment created and dependencies installed!"
echo "💡 To activate it later, run: source venv/bin/activate"

