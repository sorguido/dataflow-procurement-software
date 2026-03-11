@echo off
REM Batch script to setup virtual environment for Database Migration Tool
REM Creates venv named 'dmt' and installs dependencies

echo ======================================
echo DataFlow Migration Tool - Setup
echo ======================================
echo.

REM Check Python
echo Checking Python version...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found!
    echo Please install Python 3.10 or higher from https://www.python.org/
    pause
    exit /b 1
)
python --version
echo.

REM Create virtual environment
echo Creating virtual environment 'dmt'...
python -m venv dmt
if errorlevel 1 (
    echo ERROR: Failed to create virtual environment!
    pause
    exit /b 1
)
echo Virtual environment created successfully.
echo.

REM Activate virtual environment
echo Activating virtual environment...
call dmt\Scripts\activate.bat
if errorlevel 1 (
    echo ERROR: Failed to activate virtual environment!
    pause
    exit /b 1
)
echo Virtual environment activated.
echo.

REM Upgrade pip
echo Upgrading pip...
python -m pip install --upgrade pip --quiet
echo Pip upgraded successfully.
echo.

REM Install dependencies
echo Installing dependencies from requirements.txt...
pip install -r requirements.txt --quiet
echo Dependencies installed successfully.
echo.

REM Success message
echo ======================================
echo Setup completed successfully!
echo ======================================
echo.
echo To run the migration tool:
echo   1. Activate virtual environment:
echo      dmt\Scripts\activate.bat
echo   2. Run the tool:
echo      python Database_Migration_Tool.py
echo.
pause
