# Script: modules/Install-TinyTeX.ps1
# Purpose: Installs TinyTeX (LaTeX distribution) via Quarto.

Write-Host "Starting TinyTeX installation via Quarto..." -ForegroundColor Yellow
Write-Host "This can take a significant amount of time."

# Check if Quarto is available
if (-not (Get-Command quarto -ErrorAction SilentlyContinue)) {
    Write-Error "Quarto command not found. Cannot install TinyTeX."
    exit 1
}

Write-Host "Found Quarto: $(quarto --version)"

try {
    Write-Host "Executing: quarto install tool tinytex"
    # Suppress progress bar as it can make logs messy
    $ProgressPreference = 'SilentlyContinue'
    quarto install tool tinytex
    $exitCode = $LASTEXITCODE
    $ProgressPreference = 'Continue' # Reset preference
    
    if ($exitCode -ne 0) {
        Write-Error "TinyTeX installation command (quarto) failed with exit code $exitCode."
        exit 1
    }

    Write-Host "TinyTeX installation command executed." -ForegroundColor Green
    # Main-Installer.ps1 will re-test
    exit 0
} catch {
    Write-Error "An unexpected error occurred during TinyTeX installation: $($_.Exception.Message)"
    exit 1
}