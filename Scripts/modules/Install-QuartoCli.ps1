# Script: modules/Install-QuartoCli.ps1
# Purpose: Installs Quarto CLI using Chocolatey.

Write-Host "Starting Quarto CLI installation via Chocolatey..." -ForegroundColor Yellow

try {
    Write-Host "Executing: choco install quarto-cli --yes --force"
    choco install quarto --yes --force --timeout=3600
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Quarto CLI installation command (choco) failed with exit code $LASTEXITCODE."
        exit 1
    }

    Write-Host "Quarto CLI installation command executed." -ForegroundColor Green
    # Main-Installer.ps1 will refresh path and re-test
    exit 0
} catch {
    Write-Error "An unexpected error occurred during Quarto CLI installation: $($_.Exception.Message)"
    exit 1
}