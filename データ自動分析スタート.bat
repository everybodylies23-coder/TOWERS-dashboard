@echo off
cd /d "%~dp0"
echo ==============================================
echo  Starting Pachislot Auto Analyzer...
echo ==============================================
echo.
echo Checking Python libraries...
python -m pip install -q openpyxl google-generativeai beautifulsoup4
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Failed to run pip install. Please verify Python installation.
    echo Make sure "Add python.exe to PATH" was checked during Python setup.
    echo.
    pause
    exit /b
)

echo.
echo Running analysis script...
python auto_analyzer.py
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Analyzer script failed.
    echo.
    pause
    exit /b
)

echo.
echo Process finished successfully!
pause
