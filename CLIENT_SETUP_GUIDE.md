# Client Setup Guide - Course Management System

## 📋 What You Need to Get Started

This guide will help you set up the Course Management System on your computer, even if you don't have Python or any programming experience.

---

## 🚀 Option 1: Automatic Setup (Recommended - Easiest)

### For Windows Users:

1. **Download the application folder** to your computer
2. **Right-click** on `setup_and_run.ps1` file
3. **Select "Run with PowerShell"**
4. **Wait** for the automatic installation to complete
5. **The application will open automatically** in your web browser

**That's it!** The system will automatically:

- Check if Python is installed
- Install Python if needed
- Install all required libraries
- Start the application

---

## 🛠️ Option 2: Manual Setup (If Option 1 doesn't work)

### Step 1: Install Python

1. **Go to** [https://www.python.org/downloads/](https://www.python.org/downloads/)
2. **Click** "Download Python" (latest version)
3. **Run the installer** and **IMPORTANT**: Check "Add Python to PATH" during installation
4. **Restart your computer** after installation

### Step 2: Install Required Libraries

1. **Open Command Prompt** (Press Windows key + R, type `cmd`, press Enter)
2. **Navigate to the application folder**:
   ```
   cd "C:\path\to\your\application\folder"
   ```
3. **Install libraries**:
   ```
   pip install -r requirements.txt
   ```

### Step 3: Run the Application

1. **In the same Command Prompt**, type:
   ```
   streamlit run app.py
   ```
2. **The application will open** in your web browser automatically

---

## 🔧 Alternative Quick Start Methods

### Method A: Using Batch Files (Windows)

1. **Double-click** `quick_start.bat`
2. **Follow any on-screen instructions**

### Method B: Using PowerShell Script

1. **Right-click** `setup_and_run.ps1`
2. **Select "Run with PowerShell"**

---

## ❗ Troubleshooting Common Issues

### Python Not Found Error

**Problem**: "python is not recognized as an internal or external command"
**Solution**:

1. Reinstall Python and check "Add Python to PATH"
2. Restart your computer
3. Try again

### Permission Denied Error

**Problem**: Cannot run PowerShell scripts
**Solution**:

1. Open PowerShell as Administrator
2. Run: `Set-ExecutionPolicy RemoteSigned`
3. Type `Y` and press Enter
4. Try running the script again

### Library Installation Fails

**Problem**: pip install command fails
**Solution**:

1. Update pip: `python -m pip install --upgrade pip`
2. Try installing libraries one by one:
   ```
   pip install streamlit
   pip install pandas
   pip install openpyxl
   pip install python-docx
   ```

### Application Won't Start

**Problem**: Streamlit command not found
**Solution**:

1. Try: `python -m streamlit run app.py`
2. Or: `py -m streamlit run app.py`

---

## 💻 System Requirements

### Minimum Requirements:

- **Operating System**: Windows 10 or newer
- **RAM**: 4GB (8GB recommended)
- **Storage**: 500MB free space
- **Internet**: Required for initial setup

### Software Requirements:

- **Python 3.8 or newer** (will be installed automatically)
- **Microsoft Office or Office viewer** (for opening generated Word documents)
- **Modern web browser** (Chrome, Firefox, Edge)

---

## 📞 Getting Help

### If you encounter any issues:

1. **Check Internet Connection**: Make sure you're connected to the internet
2. **Restart Computer**: Sometimes a restart solves installation issues
3. **Run as Administrator**: Try running the setup files as administrator
4. **Contact Support**: Send us:
   - Screenshot of any error messages
   - Your Windows version
   - What you were trying to do when the error occurred

### Contact Information:

- **Email**: [your-support-email@domain.com]
- **Phone**: [your-support-phone]

---

## 🎯 Quick Reference

### To Start the Application:

1. **Double-click** `setup_and_run.ps1` (first time)
2. **Or double-click** `quick_start.bat` (subsequent uses)
3. **Or open Command Prompt** and run `streamlit run app.py`

### Application Will Open At:

- **Local URL**: http://localhost:8501
- **Network URL**: http://your-ip-address:8501

### To Stop the Application:

- **Close the browser tab**
- **Press Ctrl+C** in the Command Prompt window

---

## 📚 What This Application Does

- **Upload Excel Files**: Import your course data
- **Filter by Month**: View courses for specific months
- **Generate Documents**: Create Word documents automatically
- **Download Results**: Get your formatted documents instantly

---

**Note**: This guide assumes Windows operating system. For Mac or Linux users, please contact support for specific instructions.
