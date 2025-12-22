@echo off
REM Start a local static server for the kaoqin-management workspace and open browser.
REM Adjust PYTHON_CMD below to `python3` if your system uses that.
set PYTHON_CMD=python
cd C:\workspace\kaoqin-management
nstart "Kaoqin Server" cmd /k "%PYTHON_CMD% -m http.server 8000"
timeout /t 1 >nul
start "" "http://localhost:8000"
exit
