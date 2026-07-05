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

REM Pin BOTH the installer version and its official SHA-256. When you bump the
REM version, update PY_SHA256 to the matching hash from python.org's release page
REM (https://www.python.org/downloads/release/python-XXXX/ -> "Files" table).
set "PY_VER=3.12.10"
set "PY_SHA256=67b5b7b0b4 b8f5b6b4a... REPLACE_WITH_OFFICIAL_SHA256"
set "INSTALLER=%TEMP%\python_installer.exe"
powershell -NoProfile -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/%PY_VER%/python-%PY_VER%-amd64.exe' -OutFile '%INSTALLER%' -UseBasicParsing"

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

REM --- Verify the download's integrity before executing it -------------------
REM Skips verification only if PY_SHA256 is still the placeholder, so the tool
REM keeps working out of the box; fill in the real hash to enforce it.
echo "%PY_SHA256%" | findstr /C:"REPLACE_WITH_OFFICIAL_SHA256" >nul
if errorlevel 1 (
  for /f "skip=1 delims=" %%H in ('powershell -NoProfile -Command "(Get-FileHash -Algorithm SHA256 '%INSTALLER%').Hash"') do set "GOT_SHA=%%H"
  if /i not "%GOT_SHA%"=="%PY_SHA256%" (
    echo.
    echo [ERROR] Python installer checksum mismatch -- refusing to run it.
    echo         Expected: %PY_SHA256%
    echo         Got:      %GOT_SHA%
    echo         Delete "%INSTALLER%" and try again, or install Python manually.
    del "%INSTALLER%" >nul 2>nul
    pause
    exit /b 1
  )
  echo [OK] Installer checksum verified.
) else (
  echo [WARN] PY_SHA256 not set -- skipping integrity check. Set it to enforce.
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
