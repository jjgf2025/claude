@echo off
title DTI Report - SETUP
color 0A
echo.
echo  ============================================
echo   DTI Portal Report Downloader - SETUP
echo  ============================================
echo.

REM Provjeri da li Python postoji
python --version >nul 2>&1
if errorlevel 1 (
    echo [GRESKA] Python nije instaliran!
    echo.
    echo  Idi na https://www.python.org/downloads/
    echo  Skini Python 3.11+ i instaliraj ga.
    echo  VAZNO: Cekiraj "Add Python to PATH" tokom instalacije!
    echo.
    pause
    exit /b 1
)

echo [OK] Python je pronadjen.
echo.
echo [INFO] Instaliranje potrebnih biblioteka...
pip install selenium openpyxl webdriver-manager -q

echo.
echo [INFO] Pravljenje foldera za reporte...
if not exist "%USERPROFILE%\Desktop\DTI_Reports" mkdir "%USERPROFILE%\Desktop\DTI_Reports"

echo.
echo  ============================================
echo   SLEDECI KORAK - OperaDriver
echo  ============================================
echo.
echo  1. Idi na:
echo     https://github.com/operasoftware/operachromiumdriver/releases
echo.
echo  2. Skini verziju koja odgovara tvojoj Operi:
echo     (Provjeri Opera verziju: Opera meni - About Opera)
echo.
echo  3. Raspakuj i stavi 'operadriver.exe' u ovaj folder:
echo     %~dp0
echo.
echo  ============================================
echo   SETUP GOTOV!
echo  ============================================
echo.
echo  Sada dvaput klikni na:  DTI_Report_Downloader.vbs
echo  (ili desni klik - Posalji na - Desktop (shortcut))
echo.
pause
