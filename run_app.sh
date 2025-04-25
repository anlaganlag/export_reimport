#!/bin/bash

# Make the script exit on any error
set -e

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "Python is not installed. Please install Python 3.8 or later."
    echo "For Mac, use: brew install python3"
    echo "For Linux, use: sudo apt-get install python3"
    exit 1
fi

# Check if virtual environment exists
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
fi

# Activate virtual environment
echo "Activating virtual environment..."
source venv/bin/activate

# Install or upgrade pip
echo "Upgrading pip..."
python -m pip install --upgrade pip

# Install requirements
echo "Installing requirements..."
pip install -r requirements.txt

# Run the Streamlit app
echo "Starting Streamlit app..."
echo "Please wait for your browser to open automatically..."
streamlit run app.py

# Note: The virtual environment will be deactivated when you close the terminal
# To deactivate manually, run: deactivate 