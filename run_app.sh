#!/bin/bash

# Make the script exit on any error
set -e

# Pull latest code from git repository using absolute path
echo "拉取最新代码..."
cd /Users/ciiber/Documents/code/export_reimport
git pull

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

# Check if port 8501 is in use and kill the process if needed
echo "检查端口8501是否被占用..."
if lsof -i :8501 &> /dev/null; then
    echo "端口8501已被占用，正在终止占用进程..."
    lsof -ti :8501 | xargs kill -9
    echo "端口8501已释放"
fi

# Run the Streamlit app on port 8501
echo "Starting Streamlit app on port 8501..."
echo "Please wait for your browser to open automatically..."
streamlit run app.py --server.port 8501

# Note: The virtual environment will be deactivated when you close the terminal
# To deactivate manually, run: deactivate