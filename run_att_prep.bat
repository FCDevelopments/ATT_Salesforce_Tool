@echo off
setlocal enableextensions
REM ==========================================================================
REM  ATT to Salesforce prep RUNNER
REM  This is what the weekly scheduled task launches. It can also be
REM  double-clicked to run the merge right now.
REM
REM  It looks for the two downloaded files in the "input" folder next to this
REM  file, and writes the result CSVs into the "output" folder.
REM
REM  When launched by the scheduled task it is called as:  run_att_prep.bat scheduled
REM  (that suppresses the prompts so an unattended run never hangs).
REM ==========================================================================

set "HERE=%~dp0"
set "SCRIPT=%HERE%att_salesforce_prep.py"
set "INPUT=%HERE%input"
set "OUTPUT=%HERE%output"
set "LOG=%HERE%last_run_log.txt"

set "INTERACTIVE=1"
if /i "%~1"=="scheduled" set "INTERACTIVE=0"

REM --- Find a Python interpreter (prefer the py launcher) --------------------
set "PYEXE="
where py >nul 2>nul && set "PYEXE=py"
if not defined PYEXE where python >nul 2>nul && set "PYEXE=python"

if not defined PYEXE (
  echo [%DATE% %TIME%] ERROR: Python is not installed or not on PATH.> "%LOG%"
  echo Install Python 3 from https://www.python.org/downloads/ ^(check "Add to PATH"^), then run Setup again.>> "%LOG%"
  echo.
  echo Python was not found. See last_run_log.txt for details.
  if "%INTERACTIVE%"=="1" ( echo Press any key to close... & pause >nul )
  exit /b 1
)

REM --- Make sure the required packages are present --------------------------
%PYEXE% -c "import pandas, openpyxl" 1>nul 2>nul
if errorlevel 1 (
  echo Installing required Python packages ^(one-time^): pandas, openpyxl ...
  %PYEXE% -m pip install --user --quiet pandas openpyxl
)

REM --- Run the merge --------------------------------------------------------
echo [%DATE% %TIME%] Starting ATT to Salesforce prep > "%LOG%"
%PYEXE% "%SCRIPT%" --input-dir "%INPUT%" --output-dir "%OUTPUT%" >> "%LOG%" 2>&1
set "RC=%ERRORLEVEL%"
echo [%DATE% %TIME%] Finished with exit code %RC% >> "%LOG%"

REM Show the log when run by hand (a scheduled run just exits quietly).
if "%INTERACTIVE%"=="1" (
  type "%LOG%"
  echo.
  if "%RC%"=="0" (
    echo Done. The upload CSV is in the "output" folder. Full log: last_run_log.txt
  ) else (
    echo Something needs attention - read the message above. Full log: last_run_log.txt
  )
  echo Press any key to close...
  pause >nul
)
exit /b %RC%
