# REQUIRES: Run this script in PowerShell as Administrator.

Write-Host "--- Starting Comprehensive All-in-One Installation Script ---" -ForegroundColor Yellow
Write-Host "IMPORTANT: This script MUST be run as Administrator."
Write-Host "This version installs a broader set of common Python and R packages, and VSCode."
Read-Host -Prompt "Press Enter to continue if you are running as Administrator, or Ctrl+C to cancel"

Function Run-Command {
    param(
        [string]$Command,
        [string]$ErrorMessage
    )
    Write-Host "Executing: $Command"
    Invoke-Expression $Command
    if ($LASTEXITCODE -ne 0) {
        Write-Error "$ErrorMessage (Exit Code: $LASTEXITCODE). Please check output above. Script will attempt to continue but may fail."
    } else {
        Write-Host "Command successful."
    }
}

Function Refresh-Path {
    Write-Host "Attempting to refresh PATH environment variable for current session..."
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    Import-Module "$($env:ProgramData)\chocolatey\helpers\chocolateyProfile.psm1" -ErrorAction SilentlyContinue
    Write-Host "PATH refresh attempted. New commands might now be available."
}

# 1. Install Chocolatey (if not present)
Write-Host "`n--- Step 1/9: Installing Chocolatey (if not present) ---" -ForegroundColor Cyan
if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
    Run-Command -Command "Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))" -ErrorMessage "Chocolatey installation failed."
    Refresh-Path
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        Write-Warning "Chocolatey installed, but 'choco' command may not be available yet in this session. If subsequent steps fail, try re-running the script in a new Administrator PowerShell window."
    }
} else {
    Write-Host "Chocolatey already installed."
}

# 2. Install Python
Write-Host "`n--- Step 2/9: Installing Python ---" -ForegroundColor Cyan
Run-Command -Command "choco install python --yes --force --request-timeout=3600" -ErrorMessage "Python installation failed."
Refresh-Path

# 3. Install R
Write-Host "`n--- Step 3/9: Installing R ---" -ForegroundColor Cyan
Run-Command -Command "choco install r.project --yes --force --request-timeout=3600" -ErrorMessage "R installation failed."
Refresh-Path

# 4. Install Quarto CLI
Write-Host "`n--- Step 4/9: Installing Quarto CLI ---" -ForegroundColor Cyan
Run-Command -Command "choco install quarto-cli --yes --force --request-timeout=3600" -ErrorMessage "Quarto CLI installation failed."
Refresh-Path

# 5. Install Visual Studio Code (VSCode)
Write-Host "`n--- Step 5/9: Installing Visual Studio Code ---" -ForegroundColor Cyan
Run-Command -Command "choco install vscode --yes --force --request-timeout=3600" -ErrorMessage "Visual Studio Code installation failed."

# 6. Install Python Packages (Expanded List)
Write-Host "`n--- Step 6/9: Installing Python Packages (Expanded List) ---" -ForegroundColor Cyan
$pythonPackages = "pandas pywin32 numpy scipy matplotlib seaborn openpyxl requests"
Run-Command -Command "pip install $pythonPackages" -ErrorMessage "Python package installation (pip) failed."

# 7. Install R Packages (Expanded List - this may take a significant amount of time)
Write-Host "`n--- Step 7/9: Installing R Packages (Expanded List) ---" -ForegroundColor Cyan
Write-Host "NOTE: R package installation can take a considerable amount of time, especially for 'tidyverse'."
$rPackagesToInstall = "c('tidyverse', 'fmsb', 'scales', 'rmarkdown', 'knitr', 'openxlsx', 'readxl')"
Run-Command -Command "Rscript -e ""install.packages($rPackagesToInstall, repos='https://cloud.r-project.org/', Ncpus = max(1, parallel::detectCores() - 1))""" -ErrorMessage "R package installation (Rscript) failed."

# 8. Install TinyTeX
Write-Host "`n--- Step 8/9: Installing TinyTeX (LaTeX for Quarto) ---" -ForegroundColor Cyan
Run-Command -Command "quarto install tinytex" -ErrorMessage "TinyTeX installation (quarto) failed."

# 9. Install Quarto Extension
Write-Host "`n--- Step 9/9: Installing Quarto Extension ---" -ForegroundColor Cyan
Run-Command -Command "quarto install extension nmfs-opensci/quarto_titlepages --no-prompt" -ErrorMessage "Quarto extension installation failed."

Write-Host "`n--- All Installation Steps Attempted ---" -ForegroundColor Yellow
Write-Host "If any 'command not found' errors occurred for choco, pip, Rscript, or quarto, it's likely due to PATH environment variable updates not fully propagating in this single session."
Write-Host "In such cases, open a NEW Administrator PowerShell window; the commands should then be available."
Write-Host "Manual steps still potentially required: Install custom font (QTDublinIrish.otf) and ensure Outlook is configured."
Read-Host -Prompt "Script finished. Press Enter to exit."