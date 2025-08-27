#!/bin/bash

echo "Installing Streamlit Training Courses Management System..."
echo

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    if ! command -v python &> /dev/null; then
        echo "Python is not installed or not in PATH."
        echo "Please install Python from https://python.org"
        exit 1
    else
        PYTHON_CMD="python"
    fi
else
    PYTHON_CMD="python3"
fi

echo "Python found. Installing requirements..."
echo

# Install requirements
$PYTHON_CMD -m pip install -r requirements.txt

if [ $? -ne 0 ]; then
    echo "Failed to install requirements."
    exit 1
fi

echo
echo "Installation completed successfully!"
echo
echo "To run the application:"
echo "1. Open terminal"
echo "2. Navigate to this directory"
echo "3. Run: streamlit run app.py"
echo
echo "Sample data files are available in the sample_data folder."
echo
