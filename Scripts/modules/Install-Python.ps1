# Script: modules/Install-Python.ps1
# Purpose: Installs Python (which includes Pip) using Chocolatey.

Write-Host "Starting Python installation via Chocolatey..." -ForegroundColor Yellow

try {
    Write-Host "Executing: choco install python --yes --force"
    choco install python --yes --force --timeout=3600 # Increased timeout
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Python installation command (choco) failed with exit code $LASTEXITCODE."
        exit 1
    }

    Write-Host "Python installation command executed." -ForegroundColor Green
    # Main-Installer.ps1 will refresh path and re-test for python and pip
    exit 0
} catch {
    Write-Error "An unexpected error occurred during Python installation: $($_.Exception.Message)"
    exit 1
}