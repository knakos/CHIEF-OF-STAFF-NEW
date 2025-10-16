@echo off
title Outlook Inbox Reader - GUI

REM Check if pywin32 is installed
python -c "import win32com.client" 2>nul
if errorlevel 1 (
    echo Installing required libraries...
    pip install -r requirements.txt
    echo.
)

REM Check if customtkinter is installed
python -c "import customtkinter" 2>nul
if errorlevel 1 (
    echo Installing required libraries...
    pip install -r requirements.txt
    echo.
)

REM Run the GUI application
python app.py

REM Keep window open if there's an error
if errorlevel 1 (
    echo.
    echo An error occurred. Press any key to exit...
    pause >nul
)
