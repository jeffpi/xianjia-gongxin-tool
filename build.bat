@echo off
cd /d "%~dp0"

echo ========================================
echo  MIIT Data and Voice Tracking Tool - Build
echo ========================================
echo.

echo [1/6] Checking Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python not found. Please install Python 3.8+ first.
    pause
    exit /b 1
)
python --version
echo.

echo [2/6] Creating clean virtual environment...
if exist _build_venv rmdir /s /q _build_venv
python -m venv _build_venv
if %errorlevel% neq 0 (
    echo ERROR: Failed to create venv.
    pause
    exit /b 1
)
call _build_venv\Scripts\activate.bat
echo.

echo [3/6] Installing dependencies (clean venv)...
python -m pip install --upgrade pip -q
python -m pip install python-minifier PyInstaller PySide6 pandas numpy openpyxl python-calamine xlsxwriter pyxlsb -q
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies.
    pause
    exit /b 1
)
echo.

echo [4/6] Preparing and obfuscating source...
copy /Y "MIIT_Source.py" MIIT_DataVoiceTool.py >nul
if %errorlevel% neq 0 (
    echo ERROR: Source file MIIT_Source.py not found.
    pause
    exit /b 1
)
if exist MIIT_DataVoiceTool_obf.py del MIIT_DataVoiceTool_obf.py
python -m python_minifier --remove-literal-statements --rename-globals MIIT_DataVoiceTool.py -o MIIT_DataVoiceTool_obf.py
if %errorlevel% neq 0 (
    echo ERROR: Python-Minifier obfuscation failed.
    pause
    exit /b 1
)
echo.

echo [5/6] Building exe with PyInstaller...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
python -m PyInstaller --onefile --windowed --name "MIIT_DataVoiceTool" ^
    --add-data "hubei_monitor_host.txt;." ^
    --hidden-import PySide6 --hidden-import pandas --hidden-import numpy ^
    --hidden-import openpyxl --hidden-import xlsxwriter ^
    --noconfirm ^
    MIIT_DataVoiceTool_obf.py
if %errorlevel% neq 0 (
    echo ERROR: PyInstaller build failed.
    pause
    exit /b 1
)
echo.

echo [6/6] Cleaning up...
rmdir /s /q build
rmdir /s /q _build_venv
del MIIT_DataVoiceTool.py MIIT_DataVoiceTool_obf.py MIIT_DataVoiceTool.spec 2>nul
echo.

echo ========================================
echo  Build Complete!
echo  Output: dist\MIIT_DataVoiceTool.exe
echo ========================================
pause
