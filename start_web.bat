@echo off
cd /d "%~dp0"
echo.
echo  Meesho Web UI
echo  =============
echo  Starting server — keep this window OPEN while using the browser.
echo.
python web_app.py
if errorlevel 1 (
    echo.
    echo  Server exited with an error.
    pause
)
