@echo off
setlocal enableextensions
REM ==========================================================================
REM  ONE-TIME SETUP
REM  Double-click this file once to schedule the weekly run.
REM  It creates a Windows task that runs every MONDAY at 8:30 AM.
REM  No admin rights needed. (You must be logged in for it to run.)
REM ==========================================================================

set "HERE=%~dp0"
set "RUNNER=%HERE%run_att_prep.bat"
set "TASKNAME=ATT Salesforce Weekly Prep"

echo Creating the weekly scheduled task...
echo   Name : %TASKNAME%
echo   When : every Monday at 8:30 AM
echo   Runs : %RUNNER%
echo.

schtasks /create /tn "%TASKNAME%" /tr "\"%RUNNER%\" scheduled" /sc weekly /d MON /st 08:30 /f

if errorlevel 1 (
  echo.
  echo [PROBLEM] The task could not be created.
  echo Try again by RIGHT-CLICKING this file and choosing "Run as administrator".
  echo.
  pause
  exit /b 1
)

echo.
echo ==========================================================================
echo SUCCESS! The weekly task is set up.
echo.
echo Every Monday at 8:30 AM it will process whatever you have placed in the
echo "input" folder and write the upload CSV to the "output" folder.
echo.
echo REMINDER: each week, download the two reports and drop them in the
echo "input" folder BEFORE Monday morning.
echo ==========================================================================
echo.
echo Press any key to close...
pause >nul
