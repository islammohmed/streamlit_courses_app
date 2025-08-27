# PowerShell script to setup and run the Streamlit application
Write-Host "=== نظام إدارة الدورات التدريبية ===" -ForegroundColor Green
Write-Host "Setting up and running the Training Courses Management System..." -ForegroundColor Yellow

# Check if Python is installed
try {
    $pythonVersion = py --version 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Python is installed: $pythonVersion" -ForegroundColor Green
        $pythonCmd = "py"
    } else {
        throw "py command failed"
    }
} catch {
    try {
        $pythonVersion = python --version 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Python is installed: $pythonVersion" -ForegroundColor Green
            $pythonCmd = "python"
        } else {
            throw "python command failed"
        }
    } catch {
        Write-Host "Python is not installed or not in PATH." -ForegroundColor Red
        Write-Host "Please install Python from https://python.org" -ForegroundColor Yellow
        Write-Host "Make sure to check 'Add Python to PATH' during installation." -ForegroundColor Yellow
        Read-Host "Press Enter to exit"
        exit 1
    }
}

# Check if pip is available
try {
    & $pythonCmd -m pip --version | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "pip not found"
    }
    Write-Host "pip is available" -ForegroundColor Green
} catch {
    Write-Host "pip is not available. Please reinstall Python with pip." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Install requirements
Write-Host "Installing required packages..." -ForegroundColor Yellow
try {
    & $pythonCmd -m pip install -r requirements.txt
    if ($LASTEXITCODE -ne 0) {
        throw "Failed to install requirements"
    }
    Write-Host "All packages installed successfully!" -ForegroundColor Green
} catch {
    Write-Host "Failed to install requirements. Check your internet connection." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Check if data files exist
$sampleDataDir = "sample_data"
$excelFiles = Get-ChildItem -Path $sampleDataDir -Filter "*.xlsx" -ErrorAction SilentlyContinue
$wordFiles = Get-ChildItem -Path $sampleDataDir -Filter "*.docx" -ErrorAction SilentlyContinue

if ($excelFiles.Count -gt 0) {
    Write-Host "✓ Excel data file found: $($excelFiles[0].Name)" -ForegroundColor Green
} else {
    Write-Host "⚠️  Excel data file not found in $sampleDataDir" -ForegroundColor Yellow
    Write-Host "You can upload your Excel file through the web interface." -ForegroundColor Yellow
}

if ($wordFiles.Count -gt 0) {
    Write-Host "✓ Word template file found: $($wordFiles[0].Name)" -ForegroundColor Green
} else {
    Write-Host "⚠️  Word template file not found in $sampleDataDir" -ForegroundColor Yellow
    Write-Host "You can upload your Word template through the web interface." -ForegroundColor Yellow
}

# Run the Streamlit application
Write-Host "Starting the web application..." -ForegroundColor Green
Write-Host "The application will open in your default web browser." -ForegroundColor Yellow
Write-Host "If it doesn't open automatically, go to: http://localhost:8501" -ForegroundColor Yellow
Write-Host ""
Write-Host "Press Ctrl+C to stop the application" -ForegroundColor Red
Write-Host ""

try {
    & $pythonCmd -m streamlit run app.py
} catch {
    Write-Host "Failed to start the application. Please check the error messages above." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}
