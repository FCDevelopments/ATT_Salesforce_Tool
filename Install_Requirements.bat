@echo off
setlocal enableextensions
REM ==========================================================================
REM  REQUIREMENTS INSTALLER
REM  Run this ONCE before using the tool for the first time.
REM  It installs Python 3 (if missing) and the required packages.
REM  Double-click it -- it will tell you everything it's doing.
REM ==========================================================================

echo ==========================================================================
echo   ATT Salesforce Prep -- Requirements Installer
echo ==========================================================================
echo.

REM --- Check if Python is already installed ----------------------------------
set "PYEXE="
where py    >nul 2>nul && set "PYEXE=py"
where python >nul 2>nul && if not defined PYEXE set "PYEXE=python"

if defined PYEXE (
  echo [OK] Python is already installed.
  %PYEXE% --version
  goto :install_packages
)

REM --- Python not found -- download and install it ---------------------------
echo [INFO] Python was not found on this computer.
echo        Downloading Python 3.12 from python.org...
echo        (This may take a minute depending on your internet speed.)
echo.

set "INSTALLER=%TEMP%\python_installer.exe"
powershell -NoProfile -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.9/python-3.12.9-amd64.exe' -OutFile '%INSTALLER%' -UseBasicParsing"

if not exist "%INSTALLER%" (
  echo.
  echo [ERROR] Could not download the Python installer.
  echo         Check your internet connection and try again.
  echo         Or download Python manually from https://www.python.org/downloads/
  echo         Make sure you tick "Add Python to PATH" during install, then run this
  echo         installer again.
  echo.
  pause
  exit /b 1
)

echo [INFO] Download complete. Installing Python silently...
echo        (You may see a User Account Control prompt -- click Yes.)
echo.
"%INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_launcher=1

if errorlevel 1 (
  echo.
  echo [ERROR] Python installation failed or was cancelled.
  echo         Try running this file as Administrator (right-click -> Run as administrator).
  echo.
  pause
  exit /b 1
)

REM Refresh PATH so the new Python is visible in this session
set "PYEXE=py"
where py >nul 2>nul
if errorlevel 1 set "PYEXE=python"

echo.
echo [OK] Python installed successfully.
echo.

:install_packages
REM --- Install / upgrade required Python packages ----------------------------
echo [INFO] Installing required packages: pandas and openpyxl...
echo        (First time may take about 30-60 seconds.)
echo.
%PYEXE% -m pip install --user --quiet --upgrade pandas openpyxl

if errorlevel 1 (
  echo.
  echo [ERROR] Package installation failed.
  echo         Check your internet connection and try again.
  echo.
  pause
  exit /b 1
)

REM --- Quick smoke test ------------------------------------------------------
%PYEXE% -c "import pandas, openpyxl; print('[OK] pandas', pandas.__version__, '/ openpyxl', openpyxl.__version__)"

if errorlevel 1 (
  echo.
  echo [ERROR] Packages installed but could not be imported. Please contact your IT/automation contact.
  echo.
  pause
  exit /b 1
)

echo.
echo ==========================================================================
echo   All requirements are installed. You are ready to go!
echo.
echo   NEXT STEP: Double-click "Setup_Weekly_Schedule.bat" to schedule the
echo   automatic weekly run (Monday at 8:30 AM).
echo ==========================================================================
echo.
pause
