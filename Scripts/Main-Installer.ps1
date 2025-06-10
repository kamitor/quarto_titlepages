# Main-Installer.ps1
# Orchestrates the installation of all necessary tools and programs for kamitor/quarto_titlepages.
# REQUIRES: Run this script in PowerShell as Administrator.

# --- ASCII Art Welcome ---
Clear-Host
$Art = @"
===============================================================================
        __
       /  \
      |----|       L E C T O R A A T
      |----|
       \__/        S U P P L Y   C H A I N   F I N A N C E
        ||
       /  \
      |----|       Environment Setup for Recilience Python Report Setup 
      |----|
       \__/
===============================================================================
"@
Write-Host $Art -ForegroundColor Cyan
Write-Host ("-" * 70)
Write-Host "Welcome to the kamitor/quarto_titlepages Environment Setup Script!" -ForegroundColor Yellow
Write-Host "This script will attempt to install necessary tools and configure your system."
Write-Host ("-" * 70)
Write-Host ""

# --- Pre-flight Checks ---
Write-Host "--- Performing Pre-flight System Checks ---" -ForegroundColor Magenta

# 1. Administrator Privileges Check
Write-Host "Checking for Administrator privileges..." -ForegroundColor Gray
try {
    $currentIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $currentUser = New-Object System.Security.Principal.WindowsPrincipal($currentIdentity)

    if (-Not $currentUser.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Error "ADMINISTRATOR PRIVILEGES REQUIRED: This script must be run as Administrator."
        Write-Warning "Please re-launch PowerShell using 'Run as Administrator' and try again."
        Read-Host "Press Enter to exit"
        exit 1
    }
    Write-Host "Administrator privileges confirmed." -ForegroundColor Green
} catch {
    Write-Error "Failed to verify Administrator privileges. Error: $($_.Exception.Message)"
    Read-Host "Press Enter to exit"
    exit 1
}

# 2. Operating System Check (Windows 10 or newer recommended)
Write-Host "Checking Operating System version..." -ForegroundColor Gray
$OS = Get-CimInstance Win32_OperatingSystem
Write-Host "Operating System: $($OS.Caption), Version: $($OS.Version)"
if (($OS.Version -split "\.")[0] -lt 10) {
    Write-Warning "This script is optimized for Windows 10 or newer. You are on an older version ($($OS.Caption)). Some features might not work as expected."
} else {
    Write-Host "OS version check passed (Windows 10 or newer)." -ForegroundColor Green
}

# 3. PowerShell Version Check (Recommended 5.1 or higher, ideally 7+)
Write-Host "Checking PowerShell version..." -ForegroundColor Gray
Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)"
if ($PSVersionTable.PSVersion.Major -lt 5 -or ($PSVersionTable.PSVersion.Major -eq 5 -and $PSVersionTable.PSVersion.Minor -lt 1) ) {
    Write-Warning "PowerShell 5.1 or higher is recommended. Your version is $($PSVersionTable.PSVersion)."
    Write-Warning "Consider upgrading PowerShell for best compatibility: https://aka.ms/PSWindows"
} else {
    Write-Host "PowerShell version check passed (5.1+)." -ForegroundColor Green
}

# 4. Internet Connection Check (Basic)
Write-Host "Checking for active Internet connection (pinging google.com)..." -ForegroundColor Gray
if (Test-Connection -ComputerName "google.com" -Count 1 -Quiet -ErrorAction SilentlyContinue) {
    Write-Host "Internet connection appears to be active." -ForegroundColor Green
} else {
    Write-Warning "Could not confirm an active Internet connection by pinging google.com."
    Write-Warning "An internet connection is required to download tools and packages."
    # Ask user if they want to proceed without confirmed internet
    $proceedWithoutInternet = Read-Host "Do you want to attempt to continue anyway? (yes/no)"
    if ($proceedWithoutInternet -ne 'yes') {
        Write-Error "Exiting script as per user request due to unconfirmed internet connection."
        exit 1
    }
    Write-Warning "Proceeding without confirmed internet connection at user's risk."
}
Write-Host ("-" * 70)
Write-Host ""
Read-Host "Pre-flight checks complete. Press Enter to begin the installation process..."
Write-Host ""


# --- Initial Setup (Execution Policy) ---
# This was already here, just moved slightly to group setup steps
try {
    Set-ExecutionPolicy Bypass -Scope Process -Force -ErrorAction Stop
    Write-Host "Execution policy set to Bypass for the current process." -ForegroundColor Green
} catch {
    Write-Error "Failed to set execution policy. Error: $($_.Exception.Message)"
    Read-Host "Press Enter to exit"
    exit 1
}

# --- Configuration ---
$BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$Global:MainInstallerBaseDir = $BaseDir # For sub-scripts needing access to main script's location
$ModulesDir = Join-Path -Path $BaseDir -ChildPath "modules"

# Variables that need to be globally available for sub-scripts or are central to the process
$Global:OverallSuccess = $true # Used to track if all steps succeeded
$Global:nmfsExtensionInstallLocation = $BaseDir # For nmfs-opensci extension and its test function

$ScriptScopeCloneRepoName = "kamitor_quarto_titlepages" # Used by local Test-IsKamitorRepoCloned and to build FullClonePath
$Global:CloneParentDir = Join-Path -Path $HOME -ChildPath "Documents\GitHub" # Used by Clone-KamitorRepo.ps1
$Global:FullClonePath = Join-Path -Path $Global:CloneParentDir -ChildPath $ScriptScopeCloneRepoName # Used by Clone-KamitorRepo.ps1 and Test-IsKamitorRepoCloned
$Global:RepoUrl = "https://github.com/kamitor/quarto_titlepages.git" # Used by Clone-KamitorRepo.ps1


# --- Helper Functions ---
Function Invoke-SubScript {
    param(
        [string]$SubScriptName,
        [string]$StepDescription
    )
    Write-Host "`n--- Checking/Initiating: $StepDescription ---" -ForegroundColor Cyan
    $SubScriptPath = Join-Path -Path $ModulesDir -ChildPath $SubScriptName
    if (-not (Test-Path $SubScriptPath -PathType Leaf)) { # Ensure it's a file
        Write-Error "Sub-script not found or is not a file: $SubScriptPath"
        $Global:OverallSuccess = $false
        return $false
    }

    try {
        Write-Host "Executing sub-script: $SubScriptPath"
        # Execute the sub-script. If it has an error, it should throw or exit with non-zero.
        & $SubScriptPath
        if ($LASTEXITCODE -ne 0) {
            Write-Error "$StepDescription (sub-script $SubScriptName) failed with exit code $LASTEXITCODE."
            $Global:OverallSuccess = $false
            return $false
        }
        Write-Host "$StepDescription completed successfully via $SubScriptName." -ForegroundColor Green
        return $true
    } catch {
        Write-Error "An error occurred while running $SubScriptName for $StepDescription $($_.Exception.Message)" 
        $Global:OverallSuccess = $false
        return $false
    }
}

Function Refresh-CurrentSessionPath {
    Write-Host "Attempting to refresh PATH environment variable for current session..." -ForegroundColor DarkGray
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
    # Attempt to load Chocolatey's profile helper if it exists
    $chocoProfilePath = Join-Path -Path $env:ProgramData -ChildPath "chocolatey\helpers\chocolateyProfile.psm1"
    if (Test-Path $chocoProfilePath) {
        Import-Module $chocoProfilePath -ErrorAction SilentlyContinue
    }
    Write-Host "PATH refresh attempted." -ForegroundColor DarkGray
}
# --- Test Functions (Check if tools are installed) ---

Function Test-IsChocolateyInstalled {
    Write-Host "Checking for Chocolatey..."
    if (Get-Command choco -ErrorAction SilentlyContinue) {
        Write-Host "Chocolatey found: $(choco --version)" -ForegroundColor Green
        return $true
    }
    Write-Host "Chocolatey not found." -ForegroundColor Yellow
    return $false
}

Function Test-IsGitInstalled {
    Write-Host "Checking for Git..."
    if (Get-Command git -ErrorAction SilentlyContinue) {
        Write-Host "Git found: $( (git --version | Out-String).Trim() )" -ForegroundColor Green
        return $true
    }
    Write-Host "Git not found." -ForegroundColor Yellow
    return $false
}

Function Test-IsPythonInstalled {
    Write-Host "Checking for Python..."
    if (Get-Command python -ErrorAction SilentlyContinue) {
        Write-Host "Python found: $( (python --version 2>&1 | Out-String).Trim() )" -ForegroundColor Green
        return $true
    }
    Write-Host "Python not found." -ForegroundColor Yellow
    return $false
}

Function Test-IsPipInstalled {
    Write-Host "Checking for Pip (Python Package Installer)..."
    if (Get-Command pip -ErrorAction SilentlyContinue) {
        Write-Host "Pip found: $( (pip --version | Out-String).Trim() )" -ForegroundColor Green
        return $true
    }
    Write-Host "Pip not found." -ForegroundColor Yellow
    return $false
}

Function Test-IsRInstalled {
    Write-Host "Checking for R and Rscript..."
    $rFound = Get-Command R -ErrorAction SilentlyContinue
    $rscriptFound = Get-Command Rscript -ErrorAction SilentlyContinue
    if ($rFound -and $rscriptFound) {
        Write-Host "R found: $( (R --version | Select-String -Pattern 'R version' | Out-String).Trim() )" -ForegroundColor Green
        Write-Host "Rscript found: $( (Rscript --version 2>&1 | Out-String).Trim() )" -ForegroundColor Green
        return $true
    }
    if (-not $rFound) { Write-Host "R not found." -ForegroundColor Yellow }
    if (-not $rscriptFound) { Write-Host "Rscript not found." -ForegroundColor Yellow }
    return $false
}

Function Test-IsQuartoCliInstalled {
    Write-Host "Checking for Quarto CLI..."
    if (Get-Command quarto -ErrorAction SilentlyContinue) {
        Write-Host "Quarto CLI found: $( (quarto --version | Out-String).Trim() )" -ForegroundColor Green
        return $true
    }
    Write-Host "Quarto CLI not found." -ForegroundColor Yellow
    return $false
}

Function Test-IsVSCodeInstalled {
    Write-Host "Checking for Visual Studio Code (via Chocolatey)..."
    if (Get-Command choco -ErrorAction SilentlyContinue) {
        $vscodePackageInfo = ""
        try {
            # -r for machine readable output to simplify parsing
            $vscodePackageInfo = choco list --local-only --exact --name-only --limit-output vscode -r 
        } catch {
             Write-Warning "Error checking for VSCode with choco list: $($_.Exception.Message)"
        }

        # $LASTEXITCODE for choco list is 0 even if package not found, so check output
        if ($vscodePackageInfo -match "vscode") { 
            Write-Host "VSCode (Chocolatey package 'vscode') is listed as installed by Chocolatey." -ForegroundColor Green
            if (Get-Command code -ErrorAction SilentlyContinue) {
                Write-Host "VSCode 'code' command is available on PATH." -ForegroundColor Green
            } else {
                Write-Host "VSCode 'code' command NOT found on PATH. (This might be okay if choco installed it but PATH not updated yet for this session)" -ForegroundColor Yellow
            }
            return $true
        }
    }
    Write-Host "VSCode (Chocolatey package 'vscode') not found or not listed by Chocolatey as installed." -ForegroundColor Yellow
    return $false
}

Function Test-IsTinyTeXInstalled {
    Write-Host "Checking for TinyTeX..."
    if (-not (Test-IsQuartoCliInstalled)) {
        Write-Host "Quarto CLI not installed, cannot check for TinyTeX." -ForegroundColor Yellow
        return $false
    }

    $tinytexPathUser = Join-Path $env:APPDATA "TinyTeX"
    # Path to check for pdflatex.exe based on typical TinyTeX structure and quarto check output
    # The 'quarto check' output often shows something like 'TinyTeX\bin\windows\'
    $pdflatexRelativePathSegments = @("bin", "windows", "pdflatex.exe") 
    $pdflatexFullPath = Join-Path -Path $tinytexPathUser -ChildPath $pdflatexRelativePathSegments

    if (Test-Path $pdflatexFullPath -PathType Leaf) {
        Write-Host "TinyTeX pdflatex.exe found at user path: $pdflatexFullPath" -ForegroundColor Green
        return $true
    }
    
    Write-Host "TinyTeX pdflatex.exe not found via direct path check ($pdflatexFullPath). Trying 'quarto check' for LaTeX detection..."
    $exitCode = 1 # Default to failure for the 'quarto check' part
    try {
        $OriginalProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        
        $quartoCheckOutputJson = quarto check --json 2>$null # Suppress stderr for cleaner JSON attempt
        if ($LASTEXITCODE -eq 0 -and $quartoCheckOutputJson) {
             $checkResult = $quartoCheckOutputJson | ConvertFrom-Json -ErrorAction SilentlyContinue
             if ($checkResult -and $checkResult.formats.pdf.latex) { 
                Write-Host "Quarto check (JSON) indicates a LaTeX distribution is available: $($checkResult.formats.pdf.latex)" -ForegroundColor Green
                return $true
             } elseif ($checkResult) { 
                Write-Host "Quarto check (JSON) parsed but did not confirm LaTeX in expected structure." -ForegroundColor Yellow
             } else { 
                Write-Host "Quarto check (JSON) output was empty or failed to parse. Falling back to text check." -ForegroundColor Yellow
             }
        } else {
            Write-Host "Quarto check with --json failed (ExitCode: $LASTEXITCODE) or produced no output. Falling back to text check." -ForegroundColor Yellow
        }

        # Fallback to string parsing if JSON failed or no output, or if structure wasn't as expected
        $quartoCheckOutputText = quarto check 2>&1 | Out-String
        if ($LASTEXITCODE -eq 0 -and ($quartoCheckOutputText -match "(?i)LaTeX\s*\[\u2714\]" -or $quartoCheckOutputText -match "(?i)LaTeX\s*\[OK\]" -or $quartoCheckOutputText -match "(?i)Found LaTeX")) {
             Write-Host "Quarto check (text) indicates LaTeX is available." -ForegroundColor Green
             return $true
        } else {
            Write-Host "Quarto check (text) does not confirm LaTeX (ExitCode: $LASTEXITCODE). Output (first 300 chars): $($quartoCheckOutputText | Select-Object -First 300)" -ForegroundColor Yellow
        }
        $exitCode = $LASTEXITCODE # Capture exit code from the last 'quarto check' text attempt
    } catch {
        Write-Warning "Exception during Quarto check for TinyTeX: $($_.Exception.Message)"
        $exitCode = -1 # Indicate an error from the try-catch block itself
    } finally {
        $ProgressPreference = $OriginalProgressPreference # Reset preference
    }

    if ($exitCode -ne 0) { # This condition reflects if the 'quarto check' attempts failed or didn't find LaTeX
        Write-Host "Attempt to use 'quarto check' for TinyTeX detection failed or did not find LaTeX. Last ExitCode for 'quarto check': $exitCode" -ForegroundColor Yellow
    }
    
    Write-Host "TinyTeX not definitively detected after all checks." -ForegroundColor Yellow
    return $false
}

Function Test-IsNmfsExtensionInstalled {
    param (
        [string]$LocationOfExtensionParentDir 
    )
    $extensionDirName = "nmfs-opensci" 
    $fullExtensionPath = Join-Path -Path $LocationOfExtensionParentDir -ChildPath "_extensions\$extensionDirName"
    Write-Host "Checking for '$extensionDirName' Quarto extension in '$fullExtensionPath'..."

    if (Test-Path $fullExtensionPath -PathType Container) { 
        Write-Host "'$extensionDirName' extension found at $fullExtensionPath." -ForegroundColor Green
        return $true
    }
    Write-Host "'$extensionDirName' extension NOT found in $fullExtensionPath." -ForegroundColor Yellow
    return $false
}

Function Test-IsKamitorRepoCloned {
    Write-Host "Checking if '$ScriptScopeCloneRepoName' repository is cloned to '$($Global:FullClonePath)'..."
    if (Test-Path (Join-Path -Path $Global:FullClonePath -ChildPath ".git") -PathType Container) { 
        Write-Host "Repository found at $($Global:FullClonePath)." -ForegroundColor Green
        return $true
    }
    Write-Host "Repository NOT found at $($Global:FullClonePath) (or is not a git repository)." -ForegroundColor Yellow
    return $false
}

Function Test-IsOutlookInstalled {
    Write-Host "Checking for Microsoft Outlook installation..."
    $outlookPaths = @(
        (Join-Path $env:ProgramFiles "Microsoft Office\root\Office16\OUTLOOK.EXE"),
        (Join-Path ${env:ProgramFiles(x86)} "Microsoft Office\root\Office16\OUTLOOK.EXE"),
        (Join-Path $env:ProgramFiles "Microsoft Office\Office16\OUTLOOK.EXE"), 
        (Join-Path ${env:ProgramFiles(x86)} "Microsoft Office\Office16\OUTLOOK.EXE"),
        # For the "new" Outlook (Monarch) from Microsoft Store
        (Join-Path $env:LOCALAPPDATA "Microsoft\WindowsApps\outlookforwindows.exe"),
        (Join-Path $env:ProgramW6432 "Microsoft Office\root\Office16\OUTLOOK.EXE") # Ensure 64-bit Program Files is checked if script runs as 32-bit on 64-bit OS
    )
    # Remove duplicates that might arise from env vars on different systems
    $uniqueOutlookPaths = $outlookPaths | Select-Object -Unique

    $outlookAppxPackage = Get-AppxPackage -Name "microsoft.outlookforwindows" -ErrorAction SilentlyContinue
            
    $foundPath = $null
    foreach ($path in $uniqueOutlookPaths) {
        if ($path -and (Test-Path $path -PathType Leaf)) { # Check if path is not null/empty before Test-Path
            $foundPath = $path
            break
        }
    }

    if ($foundPath) {
        Write-Host "Microsoft Outlook (Desktop/Legacy) executable found at: $foundPath" -ForegroundColor Green
        return $true
    } elseif ($outlookAppxPackage) {
        Write-Host "Microsoft Outlook (Store App 'microsoft.outlookforwindows') package found." -ForegroundColor Green
        return $true
    }

    # Fallback: Check registry for default MAPI client
    $mapiPath = "HKLM:\SOFTWARE\Clients\Mail"
    if (Test-Path $mapiPath) {
        $defaultClient = Get-ItemProperty -Path $mapiPath -Name "(Default)" -ErrorAction SilentlyContinue
        if ($defaultClient -and $defaultClient.'(Default)' -and (($defaultClient.'(Default)' -match "Outlook") -or ($defaultClient.'(Default)' -match "Microsoft Outlook"))) {
            Write-Host "Outlook appears to be the default MAPI client via registry." -ForegroundColor Green
            return $true
        }
    }
    Write-Host "Microsoft Outlook does not appear to be installed or readily detectable." -ForegroundColor Yellow
    return $false
}

# --- Orchestration Logic ---
Write-Host "`nStarting installation checks and process..."

# 1. Chocolatey
if (-not (Test-IsChocolateyInstalled)) {
    if (-not (Invoke-SubScript -SubScriptName "Install-Chocolatey.ps1" -StepDescription "Chocolatey Package Manager Installation")) {
        # Invoke-SubScript already sets $Global:OverallSuccess = $false
    }
    Refresh-CurrentSessionPath
    if (-not (Test-IsChocolateyInstalled)) {
        Write-Error "FATAL: Chocolatey installation failed or was not detected after install attempt. Cannot proceed."
        $Global:OverallSuccess = $false
    }
}
$chocoProfilePathGlobal = Join-Path -Path $env:ProgramData -ChildPath "chocolatey\helpers\chocolateyProfile.psm1"
if ($Global:OverallSuccess -and (Test-Path $chocoProfilePathGlobal)) { 
    Import-Module $chocoProfilePathGlobal -ErrorAction SilentlyContinue
}


# Proceed only if Chocolatey is available and previous steps were successful
if ($Global:OverallSuccess -and (Test-IsChocolateyInstalled)) {

    # 2. Git
    if (-not (Test-IsGitInstalled)) {
        if (Invoke-SubScript -SubScriptName "Install-Git.ps1" -StepDescription "Git Installation") {
            Refresh-CurrentSessionPath
        } 
        Test-IsGitInstalled # Re-test and display status
    }

    # 3. Python & Pip
    $pythonSuccessfullyInstalledOrPresent = Test-IsPythonInstalled
    if (-not $pythonSuccessfullyInstalledOrPresent) {
        Write-Host "Python not found. Attempting installation via sub-script."
        if (Invoke-SubScript -SubScriptName "Install-Python.ps1" -StepDescription "Python Installation") {
            Refresh-CurrentSessionPath # First refresh attempt
            $pythonExePath = (Get-Command python -ErrorAction SilentlyContinue).Source
            if ($pythonExePath) {
                Write-Host "Python executable found at: $pythonExePath after install. Attempting to update session PATH."
                $pythonDir = Split-Path $pythonExePath
                $scriptsDir = Join-Path $pythonDir "Scripts" 
                if ($env:Path -notlike "*$pythonDir*") { $env:Path = "$pythonDir;$($env:Path)" }
                if ($env:Path -notlike "*$scriptsDir*") { $env:Path = "$scriptsDir;$($env:Path)" }
                Write-Host "Updated session PATH segment for Python (may not reflect full system PATH): $($env:Path | Select-Object -First 300)..."
            }
            if (Test-IsPythonInstalled) { # Test again after explicit path manipulation
                $pythonSuccessfullyInstalledOrPresent = $true
                Write-Host "Python is now detected after installation attempt." -ForegroundColor Green
            } else {
                Write-Error "Python installation was attempted but Python is still not detected even after PATH manipulation."
                $Global:OverallSuccess = $false 
            }
        } else { # Invoke-SubScript for Python install failed
             $Global:OverallSuccess = $false 
        }
    }

    if ($pythonSuccessfullyInstalledOrPresent -and $Global:OverallSuccess) {
        if (-not (Test-IsPipInstalled)) { 
            Write-Warning "Python is installed, but Pip was not found by Get-Command. Attempting to install/ensure Pip using 'python -m ensurepip'..."
            try {
                $pythonCmd = "python" 
                $pythonResolved = Get-Command python -ErrorAction SilentlyContinue
                if ($pythonResolved) {
                    $pythonCmd = $pythonResolved.Source
                    Write-Host "Using specific python path for ensurepip: $pythonCmd" -ForegroundColor DarkGray
                } else {
                    Write-Warning "Could not resolve 'python' command with Get-Command. ensurepip might fail if 'python' is not directly on PATH."
                }

                Write-Host "Executing: & '$pythonCmd' -m ensurepip --upgrade" -ForegroundColor DarkGray
                & $pythonCmd -m ensurepip --upgrade 
                if ($LASTEXITCODE -ne 0) {
                    Write-Error "Attempt to run '$pythonCmd -m ensurepip --upgrade' failed with exit code $LASTEXITCODE."
                    if ($LASTEXITCODE -eq 9009) {
                        Write-Error "'python' command was not found (exit 9009) when trying to run ensurepip. This indicates a critical PATH issue for Python."
                    }
                } else {
                    Write-Host "'$pythonCmd -m ensurepip --upgrade' executed. Refreshing PATH and re-checking for Pip." -ForegroundColor Green
                    Refresh-CurrentSessionPath 
                    $pipExePath = (Get-Command pip -ErrorAction SilentlyContinue).Source
                    if ($pipExePath) {
                         $pipDir = Split-Path $pipExePath
                         if ($pipDir -and ($env:Path -notlike "*$pipDir*")) {
                            Write-Host "Adding Pip directory '$pipDir' to session PATH."
                            $env:Path = "$pipDir;$($env:Path)"
                         }
                    }
                    Test-IsPipInstalled 
                }
            } catch {
                 Write-Error "An error occurred while trying to run 'python -m ensurepip --upgrade': $($_.Exception.Message)"
            }
        }
    }

    # 4. R
    if ($Global:OverallSuccess -and -not (Test-IsRInstalled)) {
        if (Invoke-SubScript -SubScriptName "Install-R.ps1" -StepDescription "R Installation") {
            Refresh-CurrentSessionPath
            # Aggressive PATH update for R (experimental)
            $rInstallDirs = @(
                "$($env:ProgramFiles)\R",
                "$($env:ProgramW6432)\R" # For 32-bit PS on 64-bit OS finding 64-bit R, or vice versa
            )
            foreach ($rBaseDir in $rInstallDirs) {
                if (Test-Path $rBaseDir) {
                    $rVersionDirs = Get-ChildItem -Path $rBaseDir -Directory -ErrorAction SilentlyContinue | Where-Object {$_.Name -match "^R-\d+\.\d+\.\d+$"}
                    foreach ($rVersionDir in $rVersionDirs) {
                        $rBinPath = Join-Path $rVersionDir.FullName "bin\x64" # Assuming x64 common
                        if (Test-Path $rBinPath -and ($env:Path -notlike "*$rBinPath*")) {
                            Write-Host "Adding R bin path to session PATH: $rBinPath" -ForegroundColor DarkGray
                            $env:Path = "$rBinPath;$($env:Path)"
                        }
                        $rBinPathi386 = Join-Path $rVersionDir.FullName "bin\i386"
                         if (Test-Path $rBinPathi386 -and ($env:Path -notlike "*$rBinPathi386*")) {
                            Write-Host "Adding R i386 bin path to session PATH: $rBinPathi386" -ForegroundColor DarkGray
                            $env:Path = "$rBinPathi386;$($env:Path)"
                        }
                    }
                }
            }
        }
        Test-IsRInstalled
    }

    # 5. Quarto CLI
    if ($Global:OverallSuccess -and -not (Test-IsQuartoCliInstalled)) {
        if (Invoke-SubScript -SubScriptName "Install-QuartoCli.ps1" -StepDescription "Quarto CLI Installation") { # Sub-script should use 'quarto', not 'quarto-cli'
            Refresh-CurrentSessionPath
        }
        Test-IsQuartoCliInstalled
    }

    # 6. Visual Studio Code
    if ($Global:OverallSuccess -and -not (Test-IsVSCodeInstalled)) {
        Invoke-SubScript -SubScriptName "Install-VSCode.ps1" -StepDescription "Visual Studio Code Installation"
        Test-IsVSCodeInstalled 
    }

    # 7. Python Packages (uses requirements.txt via Install-PythonPackages.ps1)
    if ($Global:OverallSuccess -and $pythonSuccessfullyInstalledOrPresent -and (Test-IsPipInstalled)) {
        Invoke-SubScript -SubScriptName "Install-PythonPackages.ps1" -StepDescription "Python Packages Installation from requirements.txt"
    } elseif ($Global:OverallSuccess) { 
        Write-Warning "Skipping Python packages installation."
        if (-not $pythonSuccessfullyInstalledOrPresent) { Write-Warning "Reason: Python is not available." }
        elseif (-not (Test-IsPipInstalled)) { Write-Warning "Reason: Pip is not available or could not be installed/ensured." }
    }

    # 8. R Packages
    if ($Global:OverallSuccess -and (Test-IsRInstalled)) {
        Invoke-SubScript -SubScriptName "Install-RPackages.ps1" -StepDescription "R Packages Installation"
    } elseif ($Global:OverallSuccess) { Write-Warning "Skipping R packages: R/Rscript is not available." }

    # 9. TinyTeX (requires Quarto)
    if ($Global:OverallSuccess -and (Test-IsQuartoCliInstalled)) {
        if (-not (Test-IsTinyTeXInstalled)) {
            Invoke-SubScript -SubScriptName "Install-TinyTeX.ps1" -StepDescription "TinyTeX Installation (for Quarto)"
            Test-IsTinyTeXInstalled
        }
    } elseif ($Global:OverallSuccess) { Write-Warning "Skipping TinyTeX: Quarto CLI is not available." }

    # 10. 'nmfs-opensci/quarto_titlepages' Extension (requires Quarto)
    if ($Global:OverallSuccess -and (Test-IsQuartoCliInstalled)) {
        if (-not (Test-IsNmfsExtensionInstalled -LocationOfExtensionParentDir $Global:nmfsExtensionInstallLocation)) {
            Write-Host "The 'nmfs-opensci/quarto_titlepages' extension will be installed in an '_extensions' folder within: $($Global:nmfsExtensionInstallLocation)"
            Invoke-SubScript -SubScriptName "Install-NmfsQuartoExtension.ps1" -StepDescription "'nmfs-opensci/quarto_titlepages' Quarto Extension Installation"
            Test-IsNmfsExtensionInstalled -LocationOfExtensionParentDir $Global:nmfsExtensionInstallLocation
        }
    } elseif ($Global:OverallSuccess) { Write-Warning "Skipping 'nmfs-opensci/quarto_titlepages' extension: Quarto CLI is not available." }

    # 11. Clone 'kamitor/quarto_titlepages' Repository (requires Git)
    if ($Global:OverallSuccess -and (Test-IsGitInstalled)) {
        if (-not (Test-IsKamitorRepoCloned)) {
            Invoke-SubScript -SubScriptName "Clone-KamitorRepo.ps1" -StepDescription "Cloning 'kamitor/quarto_titlepages' Repository"
            Test-IsKamitorRepoCloned
        }
    } elseif ($Global:OverallSuccess) { Write-Warning "Skipping repository cloning: Git is not available." }

} elseif (-not (Test-IsChocolateyInstalled)) { 
    Write-Error "Cannot proceed with tool installations because Chocolatey is not available."
    # $Global:OverallSuccess should already be false from choco check
}

# 12. Install Custom Fonts
if ($Global:OverallSuccess) { 
    $fontDir = Join-Path -Path $Global:MainInstallerBaseDir -ChildPath "assets\fonts" # Ensure $Global:MainInstallerBaseDir is set
    if (Test-Path $fontDir -PathType Container) {
        if ((Get-ChildItem -Path $fontDir -Filter "*.otf" -ErrorAction SilentlyContinue) -or (Get-ChildItem -Path $fontDir -Filter "*.ttf" -ErrorAction SilentlyContinue)) {
            Invoke-SubScript -SubScriptName "Install-CustomFonts.ps1" -StepDescription "Custom Font Installation"
        } else {
            Write-Host "No .otf or .ttf font files found in '$fontDir'. Skipping custom font installation." -ForegroundColor DarkGray
        }
    } else {
        Write-Warning "Font directory '$fontDir' not found. Skipping custom font installation."
        Write-Warning "Create the directory (e.g., 'assets\fonts' next to this script) and place font files there if needed."
    }
}

# 13. Check for Microsoft Outlook
if ($Global:OverallSuccess) { 
    Write-Host "`n--- Checking for Microsoft Outlook ---" -ForegroundColor Cyan
    if (-not (Test-IsOutlookInstalled)) {
        Write-Warning "Microsoft Outlook was not detected. If your project requires Outlook interaction, please ensure it is installed and configured manually."
    } else {
        Write-Host "Microsoft Outlook appears to be installed. Please ensure it is configured with a mail profile if required by the project." -ForegroundColor Green
    }
}

# --- Final Summary ---
Write-Host "`n--- Installation Orchestration Attempted ---" -ForegroundColor Yellow
if ($Global:OverallSuccess) {
    Write-Host "All checked steps appear to have completed successfully or were already satisfied." -ForegroundColor Green
} else {
    Write-Warning "One or more steps may have failed or were skipped due to prior failures. Please review the output above."
}

Write-Host "`nIMPORTANT NOTES:" -ForegroundColor Yellow
Write-Host " - If any 'command not found' errors occurred for newly installed tools, despite installation attempts,"
Write-Host "   it might be due to PATH environment variable updates not fully propagating in this single session."
Write-Host "   In such cases, CLOSE THIS POWERSHELL WINDOW AND OPEN A NEW ADMINISTRATOR POWERSHELL WINDOW."
Write-Host "   The commands should then be available."
Write-Host " - The 'nmfs-opensci/quarto_titlepages' extension was attempted to be installed into an '_extensions' folder"
Write-Host "   within the directory where this script was run: $($Global:nmfsExtensionInstallLocation)\_extensions"
Write-Host " - The main project files from '$ScriptScopeCloneRepoName' should be in: $($Global:FullClonePath) (if cloning was successful)."
Write-Host "   Navigate there to use the project, e.g., 'cd ""$($Global:FullClonePath)""' and then 'quarto render yourfile.qmd'."
Write-Host ""
Write-Host "Post-installation actions:"
Write-Host " 1. Custom font installation was attempted. If 'QTDublinIrish.otf' (or others) are still not available in applications,"
Write-Host "    a system REBOOT or LOGOFF/LOGON might be necessary for all applications to recognize them."
Write-Host " 2. Microsoft Outlook installation was checked. If your project requires it, ensure Outlook is also"
Write-Host "    properly configured with a mail profile."

Read-Host -Prompt "Script finished. Press Enter to exit."