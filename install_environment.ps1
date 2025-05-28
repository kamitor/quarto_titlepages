# REQUIRES: Run this script in PowerShell as Administrator.

Write-Host "Starting minimal setup..."
Write-Host "This script will install Chocolatey (if needed), then Python, R, Quarto, required packages, and a Quarto extension."
Write-Host "Ensure you are running this PowerShell session as Administrator."
Read-Host -Prompt "Press Enter to continue if you are running as Administrator, or Ctrl+C to cancel"

# Step 1: Install Chocolatey (if not already installed)
Write-Host "[Step 1/6] Checking/Installing Chocolatey package manager..."
if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
    Write-Host "Chocolatey not found. Attempting installation..."
    Set-ExecutionPolicy Bypass -Scope Process -Force;
    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072;
    iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Chocolatey installation command failed. Please check errors. Exiting."
        exit 1
    }
    Write-Host "Chocolatey installation command executed. Refreshing environment variables for current session..."
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    Import-Module "$($env:ProgramData)\chocolatey\helpers\chocolateyProfile.psm1" -ErrorAction SilentlyContinue
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
         Write-Warning "Chocolatey might not be immediately available in this session's PATH."
         Write-Warning "If the next steps fail with 'choco' not found, please close this PowerShell window, open a new one (as Administrator), and re-run this script."
         Read-Host -Prompt "Press Enter to attempt to continue, or Ctrl+C to stop and restart PowerShell."
    }
} else {
    Write-Host "Chocolatey is already installed."
}

# Step 2: Install Python, R, and Quarto CLI using Chocolatey
Write-Host "[Step 2/6] Installing Python, R, and Quarto CLI via Chocolatey..."
Write-Host "This may take several minutes per application."

choco install python --yes --force --request-timeout=3600
if ($LASTEXITCODE -ne 0) { Write-Error "Python installation via Chocolatey failed. Exiting."; exit 1 }

choco install r.project --yes --force --request-timeout=3600
if ($LASTEXITCODE -ne 0) { Write-Error "R installation via Chocolatey failed. Exiting."; exit 1 }

choco install quarto-cli --yes --force --request-timeout=3600
if ($LASTEXITCODE -ne 0) { Write-Error "Quarto CLI installation via Chocolatey failed. Exiting."; exit 1 }

Write-Host "Core software installation commands sent. Refreshing environment variables..."
$env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")

# Step 3: Install Python Packages
Write-Host "[Step 3/6] Installing Python packages (pandas, pywin32)..."
pip install pandas pywin32
if ($LASTEXITCODE -ne 0) {
    Write-Error "Python package (pip) installation failed."
    Write-Error "This might be because 'pip' is not yet in the PATH for this session."
    Write-Error "Try opening a new PowerShell (Admin) window and running: pip install pandas pywin32"
    Write-Error "Exiting script."
    exit 1
}

# Step 4: Install R Packages
Write-Host "[Step 4/6] Installing R packages..."
Rscript -e "install.packages(c('readr', 'dplyr', 'stringr', 'tidyr', 'ggplot2', 'fmsb', 'scales'), repos='https://cloud.r-project.org/')"
if ($LASTEXITCODE -ne 0) {
    Write-Error "R package installation failed."
    Write-Error "This might be because 'Rscript' is not yet in the PATH for this session."
    Write-Error "Try opening a new PowerShell (Admin) window and running the Rscript command from Step 4 manually."
    Write-Error "Exiting script."
    exit 1
}

# Step 5: Install TinyTeX for Quarto PDF generation
Write-Host "[Step 5/6] Installing TinyTeX (LaTeX for Quarto)..."
quarto install tinytex
if ($LASTEXITCODE -ne 0) {
    Write-Error "TinyTeX installation via Quarto failed."
    Write-Error "This might be because 'quarto' is not yet in the PATH for this session."
    Write-Error "Try opening a new PowerShell (Admin) window and running: quarto install tinytex"
    Write-Error "Exiting script."
    exit 1
}

# Step 6: Install Quarto Extension
Write-Host "[Step 6/6] Installing Quarto extension 'nmfs-opensci/quarto_titlepages'..."
quarto install extension nmfs-opensci/quarto_titlepages --no-prompt
if ($LASTEXITCODE -ne 0) {
    Write-Error "Quarto extension installation failed."
    Write-Error "This might be because 'quarto' is not yet in the PATH for this session, or an issue with the extension itself."
    Write-Error "Try opening a new PowerShell (Admin) window and running: quarto install extension nmfs-opensci/quarto_titlepages --no-prompt"
    Write-Error "Exiting script."
    exit 1
}

Write-Host ""
Write-Host "--- Minimal Setup Script Completed ---"
Write-Host "All installation commands have been executed."
Write-Host "If any 'command not found' errors occurred for pip, Rscript, or quarto, it's likely due to PATH updates requiring a new PowerShell session."
Write-Host "In such cases, open a new PowerShell window (as Administrator if re-running parts of this script) and the commands should then be available."
Write-Host "Manual steps still required: Install custom font (QTDublinIrish.otf) and ensure Outlook is configured."