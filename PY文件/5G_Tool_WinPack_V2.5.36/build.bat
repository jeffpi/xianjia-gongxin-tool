@echo off
cd /d "%~dp0"
setlocal enabledelayedexpansion
title 5G User Data Voice Tool V2.5.36 - Build EXE

echo ============================================
echo   5G User Data Voice Tool V2.5.36
echo   Windows Build Script (PyArmor + PyInstaller)
echo ============================================
echo.

set PYCMD=
set SRC=5G_UserDataVoiceTool_V2.5.36.py
set APPNAME=5G_UserDataVoiceTool_V2.5.36

:: ===== Step 1: Detect Python =====
echo [1/7] Detecting Python...
python --version >nul 2>&1
if "!errorlevel!"=="0" set PYCMD=python
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
echo [2/7] Verifying source files...
if not exist "%SRC%" (
    echo [ERROR] %SRC% not found.
    echo Make sure all files are in the same folder.
    pause
    exit /b 1
)
echo Source file OK: %SRC%
echo.

:: ===== Step 3: Install dependencies =====
echo [3/7] Installing dependencies...
%PYCMD% -m pip install --upgrade pip -q 2>nul
%PYCMD% -m pip install pyarmor -q 2>nul
if errorlevel 1 (
    echo [ERROR] Failed to install PyArmor.
    pause
    exit /b 1
)
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

:: ===== Step 4: PyArmor obfuscation =====
echo [4/7] Obfuscating with PyArmor...
if exist "dist_obf" rmdir /s /q dist_obf
%PYCMD% -m pyarmor gen --output dist_obf "%SRC%"
if errorlevel 1 (
    echo [ERROR] PyArmor obfuscation failed.
    echo If trial license is exhausted, purchase a license at https://pyarmor.com
    pause
    exit /b 1
)
echo PyArmor obfuscation OK.
echo.

:: ===== Step 5: Find obfuscated file =====
echo [5/7] Locating obfuscated output...
set OBF=
for /r "dist_obf" %%f in (*.py) do (
    echo %%f | findstr /v "__init__" | findstr /v "pyarmor_runtime" >nul
    if "!errorlevel!"=="0" set OBF=%%f
)
if "%OBF%"=="" (
    echo [ERROR] No obfuscated .py file found in dist_obf.
    dir dist_obf /s /b
    pause
    exit /b 1
)
echo Obfuscated file: %OBF%
echo.

:: ===== Step 6: Build EXE with PyInstaller =====
echo [6/7] Building EXE with PyInstaller (may take 3-6 minutes)...
%PYCMD% -m PyInstaller ^
  --onedir ^
  --windowed ^
  --name="%APPNAME%" ^
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
  "%OBF%"

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

:: ===== Step 7: Report result =====
echo [7/7] Checking output...
echo.

set EXE_DIR=dist\%APPNAME%
set EXE_PATH=%EXE_DIR%\%APPNAME%.exe

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
