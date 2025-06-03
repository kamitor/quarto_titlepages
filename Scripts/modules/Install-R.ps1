# Script: modules/Install-R.ps1
# Purpose: Installs R programming language using Chocolatey.

Write-Host "Starting R installation via Chocolatey..." -ForegroundColor Yellow

try {
    Write-Host "Executing: choco install r.project --yes --force"
    choco install r.project --yes --force --timeout=7200 # Increased timeout, R can take a while
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "R installation command (choco) failed with exit code $LASTEXITCODE."
        exit 1
    }

    Write-Host "R installation command executed." -ForegroundColor Green
    # Main-Installer.ps1 will refresh path and re-test
    exit 0
} catch {
    Write-Error "An unexpected error occurred during R installation: $($_.Exception.Message)"
    exit 1
}