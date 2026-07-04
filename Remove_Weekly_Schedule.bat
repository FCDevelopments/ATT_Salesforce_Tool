@echo off
setlocal enableextensions
REM Double-click to REMOVE the weekly scheduled task (stops the automatic runs).
set "TASKNAME=ATT Salesforce Weekly Prep"

schtasks /delete /tn "%TASKNAME%" /f
if errorlevel 1 (
  echo.
  echo Could not remove the task ^(maybe it was not set up^).
) else (
  echo.
  echo The weekly task has been removed. It will no longer run automatically.
)
echo.
echo Press any key to close...
pause >nul
