@echo off
title Excel Data Processor Builder

echo ============================================================
echo   Excel Data Processor - Build Script
echo ============================================================
echo.
echo This will install dependencies and build the executable.

echo.
echo [1/4] Checking Python...
python --version
if errorlevel 1 (
    echo ERROR: Python not found!
    pause
    exit /b 1
)
echo OK - Python found
echo.

echo [2/4] Installing dependencies...
pip install pandas>=2.0.0 openpyxl>=3.1.0 pyyaml>=6.0 pyinstaller>=5.0 odfpy>=1.4.0 xlrd>=2.0.1
if errorlevel 1 (
    echo ERROR: Failed to install packages
    pause
    exit /b 1
)
echo OK - Dependencies installed
echo.

echo [3/4] Cleaning old builds...
cd ..
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo OK - Cleanup complete
echo.


echo [4/4] Building executable...
echo This takes 2-5 minutes, please wait...
python -m PyInstaller --noconfirm --onefile --windowed --name ExcelDataProcessor --specpath build --add-data "../config.yaml;." --hidden-import openpyxl --hidden-import openpyxl.cell._writer --hidden-import yaml --hidden-import odf  --hidden-import xlrd --clean gui_launcher.py

if errorlevel 1 (
    echo ERROR: Build failed
    pause
    exit /b 1
)

if not exist "dist\ExcelDataProcessor.exe" (
    echo ERROR: Executable not created
    pause
    exit /b 1
)

echo.
echo ============================================================
echo BUILD SUCCESSFUL!
echo ============================================================
echo.
echo Executable: dist\ExcelDataProcessor.exe
dir dist\ExcelDataProcessor.exe
echo.
echo Next: Copy dist\ExcelDataProcessor.exe and config.yaml
echo       to your deployment folder with Data subfolder.
echo.
pause
exit /b 0
