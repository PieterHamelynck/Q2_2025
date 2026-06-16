@echo off
cd /d "%~dp0"
python make_gallery.py
if errorlevel 1 (
    echo.
    echo Something went wrong. If Python says pandas/openpyxl is missing, run:
    echo pip install pandas openpyxl
    pause
    exit /b 1
)
start "" "gallery.html"
