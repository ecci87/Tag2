@echo off
REM Run the Image Captioning Tool server
cd /d "%~dp0"

if not exist "venv" (
    echo Creating virtual environment...
    python -m venv venv
)

call venv\Scripts\activate.bat
pip install -q -r requirements.txt
echo.
echo === Image Captioning Tool ===
echo Open http://localhost:8899 in your browser
echo.
python server.py %*
