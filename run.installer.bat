@echo off
REM This batch file runs the PowerShell installer script.

echo Locating PowerShell...
where powershell >nul 2>nul
if %errorlevel% neq 0 (
    echo PowerShell is not found in your PATH.
    echo Please ensure PowerShell is installed and accessible.
    pause
    exit /b 1
)

echo Attempting to run the PowerShell setup script (install_environment.ps1)...
echo.
echo If prompted about execution policy, you might need to allow it for this session.
echo This script will guide you through installing necessary software.
echo.

REM -ExecutionPolicy Bypass: Temporarily bypasses execution policy for this specific command.
REM -NoProfile: Speeds up PowerShell startup slightly.
REM -File: Specifies the script file to run.
REM "%~dp0": Expands to the drive and path of the current batch script, ensuring it finds the .ps1 file.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0install_environment.ps1"

echo.
echo PowerShell setup script has finished.
pause