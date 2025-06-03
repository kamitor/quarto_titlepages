# Script: modules/Install-VSCode.ps1
# Purpose: Installs Visual Studio Code using Chocolatey.

Write-Host "Starting Visual Studio Code installation via Chocolatey..." -ForegroundColor Yellow

try {
    Write-Host "Executing: choco install vscode --yes --force"
    choco install vscode --yes --force --timeout=3600 # Increased timeout
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Visual Studio Code installation command (choco) failed with exit code $LASTEXITCODE."
        exit 1
    }

    Write-Host "Visual Studio Code installation command executed." -ForegroundColor Green
    # Main-Installer.ps1 will re-test
    exit 0
} catch {
    Write-Error "An unexpected error occurred during Visual Studio Code installation: $($_.Exception.Message)"
    exit 1
}