@echo off
setlocal EnableExtensions
set "REPO_ROOT=%~dp0.."
cd /d "%REPO_ROOT%"

set "SCRIPT=%~dp0Certify-IRRBB.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    echo   "%SCRIPT%"
    echo.
    pause
    exit /b 1
)

rem Runs the Excel evidence session of docs\EXCEL_EVIDENCE.md for the
rem checked-out commit: builds the workbooks in a new bundle folder next to
rem the checkout, runs every stage, writes host.json and validates the bundle.
echo K-IRRBB Excel evidence session
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*
set "CODE=%ERRORLEVEL%"
echo.
pause
exit /b %CODE%
