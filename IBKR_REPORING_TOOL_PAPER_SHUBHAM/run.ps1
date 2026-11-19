# run.ps1 - Windows PowerShell launcher with virtual environment

$ErrorActionPreference = "Stop"

$VenvDir = ".venv"
$Script  = "Repoting_Excel.py"

# Load .env if present
if (Test-Path ".env") {
    Get-Content ".env" | ForEach-Object {
        if ($_ -match "^\s*([^#][^=]+)=(.*)$") {
            [System.Environment]::SetEnvironmentVariable($matches[1].Trim(), $matches[2].Trim(), "Process")
        }
    }
} else {
    Write-Error "Error: .env file not found. Copy .env.example to .env and fill in credentials."
    exit 1
}

# Create virtual environment if it doesn't exist
if (-not (Test-Path $VenvDir)) {
    Write-Host "Creating virtual environment..."
    python -m venv $VenvDir
}

# Activate and install dependencies
& "$VenvDir\Scripts\Activate.ps1"
python -m pip install --upgrade pip --quiet
python -m pip install -r requirements.txt --quiet

Write-Host "Running $Script..."
python $Script
