@echo off
echo Installing necessary Python packages for Slide Wizard...

REM Check if Python 3.9.13 is installed
python --version | find "Python 3.9.13" >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo Python 3.9.13 is not installed or not found in the system PATH. Please install Python 3.9.13 and rerun this script.
    exit /b 1
)

REM Install required Python packages
pip install openai
pip install python-dotenv
pip install python-pptx
pip install Pillow

echo Installation complete. You can now run the Slide Wizard application.
pause
