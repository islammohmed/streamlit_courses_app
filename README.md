# Course Management System

A Streamlit-based application for managing course data and generating Word documents from Excel templates.

## Features

- Upload and process Excel files containing course data
- Filter courses by month (January to December)
- Generate Word documents using predefined templates
- Support for Arabic content and formatting
- User-friendly web interface

## Installation & Setup

### Quick Start (Windows)

1. Double-click `setup_and_run.ps1` to automatically install dependencies and start the app
2. Or run `quick_start.bat` for a simplified startup

### Manual Installation

1. Install Python 3.8 or higher
2. Run `pip install -r requirements.txt`
3. Start the application with `streamlit run app.py`

## Usage

1. Launch the application using one of the startup scripts
2. Upload your Excel file containing course data
3. Select the desired month to filter courses
4. Click "Generate Document" to create a Word document
5. Download the generated document

## File Structure

- `app.py` - Main Streamlit application
- `config.py` - Configuration settings
- `requirements.txt` - Python dependencies
- `sample_data/` - Sample Excel and Word template files
- Various `.bat` and `.ps1` scripts for easy startup

## System Requirements

- Python 3.8+
- Windows 10/11 (recommended)
- Microsoft Office compatibility for Word documents
- 2GB RAM minimum
- Internet connection for initial setup

## Support

For technical support or questions, please contact the development team.
