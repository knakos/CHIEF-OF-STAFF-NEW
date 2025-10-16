@echo off
echo ========================================
echo Outlook Inbox Reader
echo ========================================
echo.

REM Check if pywin32 is installed, if not install it
python -c "import win32com.client" 2>nul
if errorlevel 1 (
    echo Installing required library...
    pip install pywin32
    echo.
)

REM Run the Outlook reader script
python read_outlook_inbox.py

echo.
pause
