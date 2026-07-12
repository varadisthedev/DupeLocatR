@echo off
REM Run_DupeLocatr.bat
REM Double-click this to launch DupeLocatr.ps1 against the folder this
REM .bat file is sitting in (place both files together in the folder
REM you want scanned, or in a parent folder above it).

setlocal
set "SCRIPT_DIR=%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%DupeLocatr.ps1" -Path "%CD%"

echo.
pause
