@echo off
cd /d "%~dp0"

echo ========================================
echo  MIIT Data & Voice Tracking Tool - Build
echo ========================================
echo.

echo [1/5] Checking Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python not found. Please install Python 3.8+ first.
    pause
    exit /b 1
)
python --version
echo.

echo [2/5] Installing/updating dependencies...
python -m pip install pyarmor PySide6 pandas numpy openpyxl python-calamine xlsxwriter pyxlsb -q
if %errorlevel% neq 0 (
    echo ERROR: Failed to install dependencies.
    pause
    exit /b 1
)
echo.

echo [3/5] Preparing source file...
if exist MIIT_DataVoiceTool.py del MIIT_DataVoiceTool.py
copy /Y "工信部数据、语音跟踪统计工具_V2.5.38_CC_20260722-1555.py" MIIT_DataVoiceTool.py
if %errorlevel% neq 0 (
    echo ERROR: Source file not found.
    pause
    exit /b 1
)
echo.

echo [4/5] Obfuscating with PyArmor...
if exist dist_obf rmdir /s /q dist_obf
python -m pyarmor gen --output dist_obf MIIT_DataVoiceTool.py
if %errorlevel% neq 0 (
    echo ERROR: PyArmor obfuscation failed.
    pause
    exit /b 1
)
echo.

echo [5/5] Building exe with PyInstaller...
python -m PyInstaller --onefile --windowed --name "MIIT_DataVoiceTool" ^
    --add-data "hubei_monitor_host.txt;." ^
    --hidden-import PySide6 --hidden-import pandas --hidden-import numpy ^
    --hidden-import openpyxl --hidden-import xlsxwriter ^
    dist_obf/MIIT_DataVoiceTool.py
if %errorlevel% neq 0 (
    echo ERROR: PyInstaller build failed.
    pause
    exit /b 1
)

echo.
echo ========================================
echo  Build Complete!
echo  Output: dist\MIIT_DataVoiceTool.exe
echo ========================================
pause
