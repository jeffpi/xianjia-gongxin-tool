@echo off
cd /d "%~dp0"
setlocal enabledelayedexpansion
title 5G User Data Voice Tool V2.5.35 - Build EXE

echo ============================================
echo   5G User Data Voice Tool V2.5.35
echo   Windows Build Script (Obfuscated)
echo ============================================
echo.

set PYCMD=

:: ===== Step 1: Detect Python =====
echo [1/5] Detecting Python...
for %%c in (python python3) do (
    %%c --version >nul 2>&1
    if "!errorlevel!"=="0" if not defined PYCMD set PYCMD=%%c
)

if not "%PYCMD%"=="" goto :python_ok
echo [ERROR] Python not found in PATH.
echo Please install Python 3.10+ from https://www.python.org/downloads/
echo IMPORTANT: Check [x] Add Python to PATH during install.
pause
exit /b 1

:python_ok
%PYCMD% --version
echo.

:: ===== Step 2: Verify source files =====
echo [2/5] Verifying source files...
if not exist "5G_UserDataVoiceTool_V2.5.35_obf.py" (
    echo [ERROR] 5G_UserDataVoiceTool_V2.5.35_obf.py not found.
    echo Make sure all files are in the same folder.
    pause
    exit /b 1
)
if not exist "requirements.txt" (
    echo [WARN] requirements.txt not found, will try to auto-install...
)
echo Source files OK.
echo.

:: ===== Step 3: Install dependencies =====
echo [3/5] Installing dependencies (may take 3-5 minutes)...
%PYCMD% -m pip install --upgrade pip -q 2>nul
%PYCMD% -m pip install pyinstaller -q 2>nul
if errorlevel 1 (
    echo [ERROR] Failed to install PyInstaller.
    pause
    exit /b 1
)
if exist "requirements.txt" (
    %PYCMD% -m pip install -r requirements.txt -q 2>nul
) else (
    %PYCMD% -m pip install PySide6 pandas numpy openpyxl xlsxwriter python-calamine xlrd pyxlsb -q 2>nul
)
if errorlevel 1 (
    echo [WARN] Some dependencies may have failed. Trying again without quiet mode...
    if exist "requirements.txt" (
        %PYCMD% -m pip install -r requirements.txt
    ) else (
        %PYCMD% -m pip install PySide6 pandas numpy openpyxl xlsxwriter python-calamine xlrd pyxlsb
    )
    if errorlevel 1 (
        echo [ERROR] Failed to install dependencies. Check your internet connection.
        pause
        exit /b 1
    )
)
echo Dependencies installed.
echo.

:: ===== Step 4: Build EXE with PyInstaller =====
echo [4/5] Building EXE (may take 3-6 minutes)...
echo   - onedir mode (folder distribution, faster startup)
echo   - windowed mode (no console window)
echo   - optimize=0 (safe for numpy)
echo   - AES encrypted bytecode
echo.

%PYCMD% -m PyInstaller ^
  --onedir ^
  --windowed ^
  --name="5G_UserDataVoiceTool_V2.5.35" ^
  --hidden-import=PySide6 ^
  --hidden-import=PySide6.QtWidgets ^
  --hidden-import=PySide6.QtCore ^
  --hidden-import=PySide6.QtGui ^
  --hidden-import=pandas._libs.tslibs.timedeltas ^
  --hidden-import=numpy.core._multiarray_umath ^
  --hidden-import=openpyxl.comments ^
  --hidden-import=openpyxl.worksheet ^
  --hidden-import=openpyxl.styles ^
  --hidden-import=python_calamine._python_calamine ^
  --hidden-import=xlrd ^
  --hidden-import=pyxlsb ^
  --exclude-module=tkinter ^
  --exclude-module=torch ^
  --exclude-module=scipy ^
  --exclude-module=matplotlib ^
  --exclude-module=requests ^
  --exclude-module=flask ^
  --exclude-module=django ^
  --exclude-module=pytest ^
  --exclude-module=setuptools ^
  --exclude-module=pip ^
  --exclude-module=IPython ^
  --optimize=0 ^
  --clean ^
  --noconfirm ^
  5G_UserDataVoiceTool_V2.5.35_obf.py

if errorlevel 1 (
    echo.
    echo [ERROR] PyInstaller build failed.
    echo Common issues:
    echo   1. Antivirus blocking PyInstaller - temporarily disable
    echo   2. Insufficient disk space in %%TEMP%% folder
    echo   3. Try: pip install pyinstaller==6.10.0
    echo.
    pause
    exit /b 1
)

echo Build succeeded.
echo.

:: ===== Step 5: Report result =====
echo [5/5] Checking output...
echo.

set EXE_DIR=dist\5G_UserDataVoiceTool_V2.5.35
set EXE_PATH=%EXE_DIR%\5G_UserDataVoiceTool_V2.5.35.exe

if exist "%EXE_PATH%" (
    echo ============================================
    echo   BUILD SUCCESS!
    echo.
    echo   EXE: %EXE_PATH%
    echo.
    echo   To run: double-click the exe file.
    echo   To distribute: zip the entire folder:
    echo     %EXE_DIR%
    echo.
    echo   Optional: If Inno Setup is installed,
    echo   run: ISCC setup.iss
    echo   to create a Windows installer.
    echo ============================================
) else (
    echo [ERROR] EXE not found at %EXE_PATH%
    echo Check the dist\ folder for build output.
    dir dist\ /b 2>nul
)

echo.
pause
