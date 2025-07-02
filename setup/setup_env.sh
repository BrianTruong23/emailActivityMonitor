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
pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib openpyxl

echo "✅ Virtual environment created and dependencies installed!"
echo "💡 To activate it later, run: source venv/bin/activate"

