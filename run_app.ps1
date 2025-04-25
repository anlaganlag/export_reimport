# Check if Python is installed
$pythonCmd = Get-Command python -ErrorAction SilentlyContinue
if (-not $pythonCmd) {
    Write-Host "Python is not installed. Please install Python 3.8 or later from https://www.python.org/downloads/"
    exit 1
}

# Check if virtual environment exists
if (-not (Test-Path ".\venv")) {
    Write-Host "Creating virtual environment..."
    python -m venv venv
}

# Activate virtual environment
Write-Host "Activating virtual environment..."
.\venv\Scripts\Activate.ps1

# Install or upgrade pip
Write-Host "Upgrading pip..."
python -m pip install --upgrade pip

# Install requirements
Write-Host "Installing requirements..."
pip install -r requirements.txt

# Run the Streamlit app
Write-Host "Starting Streamlit app..."
Write-Host "Please wait for your browser to open automatically..."
streamlit run app.py

# Note: The virtual environment will be deactivated when you close the terminal
# To deactivate manually, run: deactivate 