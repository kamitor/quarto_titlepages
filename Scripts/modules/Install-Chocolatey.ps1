# Script: modules/Install-Chocolatey.ps1
# Purpose: Installs the Chocolatey package manager if not already installed.

Write-Host "Starting Chocolatey Package Manager installation..." -ForegroundColor Yellow

# Check if Chocolatey is already installed (though Main-Installer.ps1 already does this,
# this sub-script might be run independently for testing)
if (Get-Command choco -ErrorAction SilentlyContinue) {
    Write-Host "Chocolatey is already installed: $(choco --version)" -ForegroundColor Green
    exit 0
}

Write-Host "Chocolatey not found. Proceeding with installation."

try {
    Write-Host "Setting execution policy to Bypass for current process (required for Chocolatey install script)..."
    Set-ExecutionPolicy Bypass -Scope Process -Force -ErrorAction Stop
    Write-Host "Execution policy set."

    Write-Host "Setting TLS 1.2 (or higher) for current session..."
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
    Write-Host "TLS protocol set."

    Write-Host "Downloading and executing Chocolatey installation script..."
    iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Chocolatey installation script execution failed with exit code $LASTEXITCODE."
        exit 1
    }

    # Verify choco is now available on PATH (might require path refresh from main script)
    if (Get-Command choco -ErrorAction SilentlyContinue) {
        Write-Host "Chocolatey installation appears successful: $(choco --version)" -ForegroundColor Green
    } else {
        Write-Warning "Chocolatey installation script ran, but 'choco' command is not yet available. A PATH refresh and possibly a new PowerShell session might be needed."
        # Main script handles PATH refresh. Exit 0 as the script itself might have run okay.
    }
    
    exit 0
} catch {
    Write-Error "An unexpected error occurred during Chocolatey installation: $($_.Exception.Message)"
    exit 1
}