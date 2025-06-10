# Script: modules/Install-Python.ps1
# Purpose: Downloads the official Python installer and guides the user through manual installation.

Write-Host "Starting Python setup: User-guided installation." -ForegroundColor Yellow

# --- Configuration ---
# You might want to parameterize the version or make it dynamic
$PythonVersion = "3.11.9" # Example: Specify a known good version. Update as needed.
# Construct a plausible download URL. Python.org URLs can change.
# For 3.11.9 it was: https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe
# It's better to go to the python.org downloads page and find the latest stable embeddable or exe.
# For simplicity here, let's assume a direct link pattern or a known good one.
# A more robust solution might scrape the downloads page or use a Chocolatey-like manifest.

# To find current versions/links: https://www.python.org/downloads/windows/
# Let's pick a specific version for reliability in the script.
# Example for Python 3.11.9 (check for latest stable if updating script)
$PythonInstallerUrl = "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-amd64.exe"
$InstallerFileName = "python-$PythonVersion-amd64.exe"
$DownloadPath = Join-Path -Path $env:TEMP -ChildPath $InstallerFileName

# --- Download Phase ---
Write-Host "Attempting to download Python $PythonVersion installer from python.org..."
Write-Host "URL: $PythonInstallerUrl"
Write-Host "Download location: $DownloadPath"

try {
    # Ensure TLS 1.2 for downloading
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri $PythonInstallerUrl -OutFile $DownloadPath -ErrorAction Stop
    Write-Host "Python installer downloaded successfully to $DownloadPath" -ForegroundColor Green
} catch {
    Write-Error "Failed to download Python installer. Error: $($_.Exception.Message)"
    Write-Warning "Please manually download the Python installer for Windows (64-bit recommended) from https://www.python.org/downloads/windows/"
    Write-Warning "Ensure you select 'Add Python to PATH' during installation."
    Read-Host "Press Enter after you have manually downloaded AND installed Python, then this script will try to continue."
    # We can't really verify manual install here easily, so we proceed based on user confirmation
    exit 0 # Exit sub-script, main script will re-test for Python
}

# --- User Installation Phase ---
Write-Host ("-"*50) -ForegroundColor Green
Write-Host "PYTHON INSTALLATION REQUIRED - USER ACTION NEEDED" -ForegroundColor Yellow
Write-Host ("-"*50) -ForegroundColor Green
Write-Host "The Python installer has been downloaded to:"
Write-Host "  $DownloadPath"
Write-Host ""
Write-Host "Please MANUALLY run this installer now." -ForegroundColor Yellow
Write-Host "IMPORTANT INSTRUCTIONS DURING INSTALLATION:" -ForegroundColor Cyan
Write-Host "  1. On the first screen of the Python installer, BE SURE TO CHECK THE BOX that says:"
Write-Host "     'Add python.exe to PATH' or 'Add Python X.Y to PATH'." 
Write-Host "     This is CRUCIAL for Python to work correctly with other tools."
Write-Host "  2. You can typically choose the default 'Install Now' option after checking the PATH box."
Write-Host "  3. Follow the on-screen prompts to complete the installation."
Write-Host ""
Write-Host "This script will wait for you to complete the Python installation." -ForegroundColor Yellow

# Open the folder containing the installer for the user
try {
    Show-WindowsExplorer -Path $DownloadPath
} catch {
    Write-Warning "Could not automatically open the download folder. Please navigate to $DownloadPath manually."
}


while ($true) {
    $confirmation = Read-Host -Prompt "Have you completed the Python installation AND ensured 'Add Python to PATH' was checked? (yes/no)"
    if ($confirmation -eq 'yes') {
        Write-Host "Thank you for confirming Python installation." -ForegroundColor Green
        break
    } elseif ($confirmation -eq 'no') {
        Write-Warning "Please complete the Python installation as instructed."
    } else {
        Write-Warning "Invalid input. Please type 'yes' or 'no'."
    }
}

# Cleanup the downloaded installer (optional)
# Remove-Item -Path $DownloadPath -Force -ErrorAction SilentlyContinue
# Write-Host "Cleaned up downloaded installer."

Write-Host "Python manual installation step complete. The main script will now re-check for Python."

exit 0