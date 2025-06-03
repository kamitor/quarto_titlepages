# Script: modules/Install-NmfsQuartoExtension.ps1
# Purpose: Installs the 'nmfs-opensci/quarto_titlepages' Quarto extension.

Write-Host "Starting 'nmfs-opensci/quarto_titlepages' Quarto extension installation..." -ForegroundColor Yellow

# Check if Quarto is available
if (-not (Get-Command quarto -ErrorAction SilentlyContinue)) {
    Write-Error "Quarto command not found. Cannot install Quarto extension."
    exit 1
}
Write-Host "Found Quarto: $(quarto --version)"

# This script expects $Global:nmfsExtensionInstallLocation to be set by Main-Installer.ps1
# to the directory where the extension should be installed (e.g., script's base directory).
# Quarto installs extensions relative to the Current Working Directory (CWD).
$TargetInstallDir = $Global:nmfsExtensionInstallLocation
if (-not ($TargetInstallDir) -or -not (Test-Path $TargetInstallDir -PathType Container)) {
    Write-Error "Target installation directory variable (\$Global:nmfsExtensionInstallLocation) is not set or directory does not exist: '$TargetInstallDir'"
    exit 1
}

$OriginalLocation = Get-Location
try {
    Write-Host "Changing working directory to: $TargetInstallDir for extension installation."
    Set-Location -Path $TargetInstallDir -ErrorAction Stop

    Write-Host "Executing: quarto install extension nmfs-opensci/quarto_titlepages --no-prompt (in $PWD)"
    quarto install extension nmfs-opensci/quarto_titlepages --no-prompt
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "'nmfs-opensci/quarto_titlepages' extension installation failed with exit code $LASTEXITCODE."
        Set-Location -Path $OriginalLocation # Ensure we go back
        exit 1
    }

    Write-Host "'nmfs-opensci/quarto_titlepages' extension installation command executed." -ForegroundColor Green
    Write-Host "Extension should be in '$TargetInstallDir\_extensions\nmfs-opensci'"
    
    Set-Location -Path $OriginalLocation
    exit 0
} catch {
    Write-Error "An unexpected error occurred during 'nmfs-opensci/quarto_titlepages' extension installation: $($_.Exception.Message)"
    if ($OriginalLocation) { Set-Location -Path $OriginalLocation }
    exit 1
}