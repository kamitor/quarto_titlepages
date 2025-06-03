# Script: modules/Install-Git.ps1
# Purpose: Installs Git using Chocolatey.

Write-Host "Starting Git installation via Chocolatey..." -ForegroundColor Yellow

try {
    Write-Host "Executing: choco install git.install -params ""/GitAndUnixToolsOnPath"" --yes --force"
    choco install git.install -params '"/GitAndUnixToolsOnPath"' --yes --force --timeout=3600 # Increased timeout
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Git installation command (choco) failed with exit code $LASTEXITCODE."
        exit 1
    }

    Write-Host "Git installation command executed." -ForegroundColor Green
    # Main-Installer.ps1 will refresh path and re-test
    exit 0
} catch {
    Write-Error "An unexpected error occurred during Git installation: $($_.Exception.Message)"
    exit 1
}