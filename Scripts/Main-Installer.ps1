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

Function Test-IsRStudioInstalled {
    Write-Host "Checking for RStudio (via Chocolatey)..."
    if (Get-Command choco -ErrorAction SilentlyContinue) {
        $rstudioPackageInfo = ""
        try {
            $rstudioPackageInfo = choco list --local-only --exact --name-only --limit-output rstudio -r 
        } catch {
             Write-Warning "Error checking for RStudio with choco list: $($_.Exception.Message)"
        }

        if ($rstudioPackageInfo -match "rstudio") { 
            Write-Host "RStudio (Chocolatey package 'rstudio') is listed as installed." -ForegroundColor Green
            # Optionally, check for the executable if you know its common path, though less reliable than choco list
            # For example: if (Test-Path "$($env:ProgramFiles)\RStudio\rstudio.exe" -PathType Leaf) { Write-Host "RStudio executable found."}
            return $true
        }
    }
    Write-Host "RStudio (Chocolatey package 'rstudio') not found or not listed by Chocolatey." -ForegroundColor Yellow
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
    # Correctly form the full path to pdflatex.exe
    $pdflatexFullPath = Join-Path -Path $tinytexPathUser -ChildPath "bin\win32\pdflatex.exe" # Common path, was 'bin\windows' before, let's try 'win32' as often seen
                                                                                         # Or use: Join-Path (Join-Path $tinytexPathUser "bin") "win32\pdflatex.exe"

    Write-Host "Checking for TinyTeX pdflatex at: $pdflatexFullPath" -ForegroundColor DarkGray
    if (Test-Path $pdflatexFullPath -PathType Leaf) {
        Write-Host "TinyTeX pdflatex.exe found at user path: $pdflatexFullPath" -ForegroundColor Green
        return $true
    } else {
        # Try the other common path variation just in case
        $pdflatexFullPathAlt = Join-Path -Path $tinytexPathUser -ChildPath "bin\windows\pdflatex.exe"
        Write-Host "Checking for TinyTeX pdflatex at alternate path: $pdflatexFullPathAlt" -ForegroundColor DarkGray
        if (Test-Path $pdflatexFullPathAlt -PathType Leaf) {
            Write-Host "TinyTeX pdflatex.exe found at user path: $pdflatexFullPathAlt" -ForegroundColor Green
            return $true
        }
    }
    
    Write-Host "TinyTeX pdflatex.exe not found via direct path checks. Trying 'quarto check' for LaTeX detection..."
    $exitCode = 1 
    try {
        $OriginalProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        
        $quartoCheckOutputJson = quarto check --json 2>$null 
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

        $quartoCheckOutputText = quarto check 2>&1 | Out-String
        if ($LASTEXITCODE -eq 0 -and ($quartoCheckOutputText -match "(?i)LaTeX\s*\[\u2714\]" -or $quartoCheckOutputText -match "(?i)LaTeX\s*\[OK\]" -or $quartoCheckOutputText -match "(?i)Found LaTeX")) {
             Write-Host "Quarto check (text) indicates LaTeX is available." -ForegroundColor Green
             return $true
        } else {
            Write-Host "Quarto check (text) does not confirm LaTeX (ExitCode: $LASTEXITCODE). Output (first 300 chars): $($quartoCheckOutputText | Select-Object -First 300)" -ForegroundColor Yellow
        }
        $exitCode = $LASTEXITCODE 
    } catch {
        Write-Warning "Exception during Quarto check for TinyTeX: $($_.Exception.Message)"
        $exitCode = -1 
    } finally {
        $ProgressPreference = $OriginalProgressPreference 
    }

    if ($exitCode -ne 0) { 
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
    Write-Host "`n--- Processing Python & Pip ---" -ForegroundColor Cyan
    $pythonSuccessfullyInstalledOrPresent = $false
    $pythonExeToUse = $null # Will hold the path to the non-stub python if found
    $attemptedStubRemovalThisRun = $false # Reset for each run of this block if script is re-entrant

    Function _Private_HandlePythonDetectionAndStub {
        param(
            [Parameter(Mandatory=$false)]
            [switch]$AfterUserInstallPrompt
        )
        $localPythonExePath = $null
        $localPythonFound = $false
        $isStubProblemCurrently = $false

        $allPythons = Get-Command python -All -ErrorAction SilentlyContinue
        
        if ($allPythons) {
            if ($AfterUserInstallPrompt) { Write-Host "Re-evaluating 'python' command(s) after user was prompted to install:" -ForegroundColor DarkGray }
            else { Write-Host "Initial 'python' command(s) found on PATH:" -ForegroundColor DarkGray }
            $allPythons | ForEach-Object { Write-Host "  - $($_.Source)" -ForegroundColor DarkGray }

            foreach ($pyInfo in $allPythons) {
                if ($pyInfo.Source -and $pyInfo.Source -notlike "*\AppData\Local\Microsoft\WindowsApps\python.exe") {
                    $localPythonExePath = $pyInfo.Source
                    break 
                }
            }
            
            if ($localPythonExePath) {
                Write-Host "Selected non-stub Python for evaluation: $localPythonExePath" -ForegroundColor DarkGreen
                $script:pythonExeToUse = $localPythonExePath
                # Direct version check for the selected Python
                if (Test-Path $script:pythonExeToUse -PathType Leaf) {
                    $output = & $script:pythonExeToUse --version 2>&1 | Out-String
                    if ($LASTEXITCODE -eq 0 -and $output -match "Python \d+\.\d+") {
                        Write-Host "Verified Python version from '$($script:pythonExeToUse)': $($output.Trim())" -ForegroundColor Green
                        $localPythonFound = $true
                    } else {
                        Write-Warning "Path $script:pythonExeToUse exists but --version failed. Exit: $LASTEXITCODE, Output: $output"
                        $script:pythonExeToUse = $null 
                    }
                } else {
                     Write-Warning "Path $script:pythonExeToUse not found or not a file."
                     $script:pythonExeToUse = $null
                }
            } elseif ($allPythons[0].Source -like "*\AppData\Local\Microsoft\WindowsApps\python.exe") {
                Write-Warning "CRITICAL: The primary 'python.exe' found is the Windows Store stub: $($allPythons[0].Source)"
                $isStubProblemCurrently = $true
                # No automatic removal, just strong warning
                Write-Warning "This stub WILL prevent proper Python operation and tool installation (pip)."
                Write-Warning "It is STRONGLY RECOMMENDED to disable 'App execution aliases' for 'python.exe' and 'python3.exe' in Windows Settings."
                Write-Warning "To do this: Search for 'Manage app execution aliases' in Windows Start, then turn them OFF."
                Write-Warning "After disabling them, CLOSE this PowerShell window and RE-RUN this script in a new Administrator PowerShell session."
                $Global:OverallSuccess = $false # Mark as failure due to persistent stub
                $script:pythonExeToUse = $allPythons[0].Source # Keep for diagnostics
            }
        } else { 
            if ($AfterUserInstallPrompt) { Write-Warning "Get-Command python still found no python.exe after user install prompt."}
            else { Write-Host "No 'python' command initially found on PATH." }
        }
        $script:pythonSuccessfullyInstalledOrPresent = $localPythonFound
        return $localPythonFound # Return if a non-stub, working Python was found
    }

    _Private_HandlePythonDetectionAndStub | Out-Null

    if (-not $pythonSuccessfullyInstalledOrPresent -and $Global:OverallSuccess) { # If no usable python and stub didn't cause hard fail
        Write-Host "Python not found or problematic. Will guide user through manual installation..."
        if (Invoke-SubScript -SubScriptName "Install-Python.ps1" -StepDescription "User-Guided Python Installation") {
            Write-Host "User has confirmed Python installation. Refreshing PATH and re-evaluating Python..." -ForegroundColor DarkGray
            Refresh-CurrentSessionPath 
            Start-Sleep -Seconds 2 # Give PATH a moment
            
            _Private_HandlePythonDetectionAndStub -AfterUserInstallPrompt:$true | Out-Null

            if ($pythonExeToUse -and (Test-Path $pythonExeToUse -PathType Leaf) -and ($pythonExeToUse -notlike "*\AppData\Local\Microsoft\WindowsApps\python.exe")) {
                Write-Host "Python executable for PATH setup post-user-install: $pythonExeToUse" -ForegroundColor DarkGreen
                $pythonDir = Split-Path $pythonExeToUse
                $scriptsDir = Join-Path $pythonDir "Scripts" 
                
                if ($pythonDir -and ($env:Path -notlike "*$(Split-Path $pythonDir -Leaf)*")) { 
                    $env:Path = "$pythonDir;$($env:Path)"
                    Write-Host "Added '$pythonDir' to session PATH." -ForegroundColor DarkGray
                }
                if ($scriptsDir -and (Test-Path $scriptsDir -PathType Container) -and ($env:Path -notlike "*$(Split-Path $scriptsDir -Leaf)*")) { 
                    $env:Path = "$scriptsDir;$($env:Path)"
                    Write-Host "Added '$scriptsDir' to session PATH." -ForegroundColor DarkGray
                }
            } elseif ($pythonExeToUse -and $pythonExeToUse -like "*\AppData\Local\Microsoft\WindowsApps\python.exe") {
                Write-Warning "Windows Store Python stub is still primary after user install. PATH modification skipped. Manual alias disabling is CRITICAL."
            } else {
                 Write-Warning "Could not locate a usable python.exe for PATH modification even after user install."
            }
            
            if (Test-IsPythonInstalled) { # This relies on Get-Command
                $pythonSuccessfullyInstalledOrPresent = $true 
                Write-Host "Python is now detected by Test-IsPythonInstalled after user installation." -ForegroundColor Green
                # Re-confirm $pythonExeToUse with the one Test-IsPythonInstalled found, if it's not a stub
                $currentFoundPython = (Get-Command Python -ErrorAction SilentlyContinue).Source
                if ($currentFoundPython -and $currentFoundPython -notlike "*\AppData\Local\Microsoft\WindowsApps\python.exe") {
                    $pythonExeToUse = $currentFoundPython
                } elseif ($currentFoundPython) { # Stub is still what Get-Command finds
                    Write-Warning "Test-IsPythonInstalled passed, but Get-Command still points to the stub: $currentFoundPython. Pip/package steps will fail."
                    $Global:OverallSuccess = $false 
                }
            } else {
                Write-Error "Python is STILL NOT DETECTED by Test-IsPythonInstalled after user install prompt and PATH refresh."
                Write-Warning "Ensure 'Add Python to PATH' was checked during manual install. A new PowerShell session might be needed."
                if ($Global:OverallSuccess) { $Global:OverallSuccess = $false }
            }
        } else { 
             Write-Error "User-guided Python installation sub-script itself reported an issue or was exited prematurely."
             $Global:OverallSuccess = $false 
        }
    }
    
    # Pip processing (mostly same as before, but relies on $pythonSuccessfullyInstalledOrPresent and $pythonExeToUse)
    if ($pythonSuccessfullyInstalledOrPresent -and $Global:OverallSuccess) {
        Write-Host "Python believed to be present. Checking for Pip..." -ForegroundColor DarkGray
        if (-not (Test-IsPipInstalled)) { 
            Write-Warning "Pip was NOT found by Test-IsPipInstalled. Attempting to ensurepip..."
            try {
                $pythonCmdForEnsurePip = $null
                if ($pythonExeToUse -and (Test-Path $pythonExeToUse -PathType Leaf) -and ($pythonExeToUse -notlike "*\AppData\Local\Microsoft\WindowsApps\python.exe")) {
                    $pythonCmdForEnsurePip = $pythonExeToUse
                } else { 
                    Write-Warning "No definitive non-stub Python identified for ensurepip. Re-checking PATH for 'python'."
                    $firstPythonOnPath = (Get-Command python -ErrorAction SilentlyContinue).Source
                    if ($firstPythonOnPath -and $firstPythonOnPath -notlike "*\AppData\Local\Microsoft\WindowsApps\python.exe") {
                        $pythonCmdForEnsurePip = $firstPythonOnPath
                    } elseif ($firstPythonOnPath) {
                        Write-Error "Cannot run ensurepip: The only 'python' found on PATH is the Windows Store stub: $firstPythonOnPath"
                        $Global:OverallSuccess = $false 
                    } else {
                        Write-Error "Cannot run ensurepip: No 'python' command found on PATH at all."
                        $Global:OverallSuccess = $false 
                    }
                }

                if ($pythonCmdForEnsurePip -and $Global:OverallSuccess) {
                    Write-Host "Using Python for ensurepip: $pythonCmdForEnsurePip" -ForegroundColor DarkGray
                    Write-Host "Executing: & '$pythonCmdForEnsurePip' -m ensurepip --upgrade" 
                    & $pythonCmdForEnsurePip -m ensurepip --upgrade 
                    if ($LASTEXITCODE -ne 0) {
                        Write-Error "Attempt to run '$pythonCmdForEnsurePip -m ensurepip --upgrade' failed with exit code $LASTEXITCODE."
                        if ($LASTEXITCODE -eq 9009) { Write-Error "Ensurepip: Python command not found (exit 9009)." }
                         $Global:OverallSuccess = $false 
                    } else {
                        Write-Host "'$pythonCmdForEnsurePip -m ensurepip --upgrade' executed. Refreshing PATH..." -ForegroundColor Green
                        Refresh-CurrentSessionPath 
                        Start-Sleep -Seconds 1
                        $pipExePathAfterEnsure = (Get-Command pip -ErrorAction SilentlyContinue).Source
                        if ($pipExePathAfterEnsure) {
                             $pipDirAfterEnsure = Split-Path $pipExePathAfterEnsure
                             if ($pipDirAfterEnsure -and (Test-Path $pipDirAfterEnsure -PathType Container) -and ($env:Path -notlike "*$(Split-Path $pipDirAfterEnsure -Leaf)*")) {
                                Write-Host "Adding Pip directory '$pipDirAfterEnsure' to session PATH." -ForegroundColor DarkGray
                                $env:Path = "$pipDirAfterEnsure;$($env:Path)"
                             }
                        }
                    }
                }
            } catch {
                 Write-Error "An exception occurred while trying to run 'python -m ensurepip --upgrade': $($_.Exception.Message)"
                 $Global:OverallSuccess = $false
            }
            
            if (Test-IsPipInstalled) {
                Write-Host "Pip is now successfully detected." -ForegroundColor Green
            } else {
                 if ($Global:OverallSuccess){ 
                    Write-Error "Pip is STILL NOT DETECTED even after ensurepip attempt."
                    Write-Warning "Python packages cannot be installed. Ensure 'Add Python to PATH' was checked. A new PowerShell session might be required."
                 }
            }
        } else {
            Write-Host "Pip was already detected." -ForegroundColor Green
        }
    } else {
        Write-Warning "Skipping Pip detection and ensurepip because Python was not successfully installed or configured due to earlier issues."
    }

        # 4. R
    Write-Host "`n--- Processing R ---" -ForegroundColor Cyan
    $Global:RScriptPath = $null # Initialize a global variable to store the path to Rscript.exe

    if (-not (Test-IsRInstalled)) {
        if (Test-IsChocolateyInstalled) {
            if (Invoke-SubScript -SubScriptName "Install-R.ps1" -StepDescription "R Installation") {
                Refresh-CurrentSessionPath # Attempt to refresh PATH
                Write-Host "R installation script completed. Attempting to locate Rscript.exe directly..." -ForegroundColor DarkGray
                
                # Try to find Rscript directly in common installation paths
                $commonRBasePaths = @(
                    "$($env:ProgramFiles)\R",
                    "$($env:ProgramW6432)\R" # For 64-bit Program Files
                ) | Select-Object -Unique
                
                foreach ($rBasePath in $commonRBasePaths) {
                    if ($rBasePath -and (Test-Path $rBasePath -PathType Container)) {
                        $rVersionDirs = Get-ChildItem -Path $rBasePath -Directory -ErrorAction SilentlyContinue | Where-Object {$_.Name -match "^R-\d+\.\d+\.\d+$"} | Sort-Object Name -Descending
                        if ($rVersionDirs) {
                            $latestRVersionDir = $rVersionDirs[0].FullName
                            $rscriptExePath = Join-Path $latestRVersionDir "bin\Rscript.exe"
                            $rExePath = Join-Path $latestRVersionDir "bin\R.exe"
                            if (Test-Path $rscriptExePath -PathType Leaf) {
                                Write-Host "Found Rscript.exe at: $rscriptExePath" -ForegroundColor Green
                                $Global:RScriptPath = $rscriptExePath
                                # Add its directory to the current session's PATH if not already there
                                $rBinDir = Split-Path $rscriptExePath
                                if ($env:PATH -notlike "*$($rBinDir)*") {
                                    Write-Host "Temporarily adding $rBinDir to session PATH." -ForegroundColor DarkGray
                                    $env:PATH = "$rBinDir;$($env:PATH)"
                                }
                                break # Found it, no need to check other base paths
                            }
                        }
                    }
                }
                if (-not $Global:RScriptPath) {
                    Write-Warning "Could not locate Rscript.exe directly after installation. PATH might not have updated yet."
                }
            } 
        } else {
            Write-Warning "Skipping R installation: Chocolatey is not available."
        }
        
        # Re-test using the updated PATH (if direct find failed) or the original PATH
        Test-IsRInstalled # This will show if Get-Command now finds it
        if ($Global:RScriptPath) {
             Write-Host "Using Rscript path for R package installation: $($Global:RScriptPath)" -ForegroundColor Yellow
        } elseif (-not (Get-Command Rscript -ErrorAction SilentlyContinue)) {
             Write-Warning "R (R.exe or Rscript.exe) not detected after installation attempt. R packages cannot be installed by this script run."
        }

    } else {
        Write-Host "R and Rscript already detected by Get-Command." -ForegroundColor Green
        $Global:RScriptPath = (Get-Command Rscript -ErrorAction SilentlyContinue).Source
         if ($Global:RScriptPath) {
             Write-Host "Using existing Rscript path: $($Global:RScriptPath)" -ForegroundColor DarkGray
         }
    }


    # 5. Quarto CLI
    Write-Host "`n--- Processing Quarto CLI ---" -ForegroundColor Cyan
    if (-not (Test-IsQuartoCliInstalled)) {
        if (Test-IsChocolateyInstalled) {
            Invoke-SubScript -SubScriptName "Install-QuartoCli.ps1" -StepDescription "Quarto CLI Installation" 
            Refresh-CurrentSessionPath
            if (-not (Test-IsQuartoCliInstalled)) {
                Write-Error "Quarto CLI not detected after installation attempt."
                $Global:OverallSuccess = $false
            }
        } else {
             Write-Warning "Skipping Quarto CLI installation: Chocolatey is not available."
             $Global:OverallSuccess = $false
        }
    } else {
        Write-Host "Quarto CLI already detected." -ForegroundColor Green
    }


    # 6. Visual Studio Code
    Write-Host "`n--- Processing Visual Studio Code ---" -ForegroundColor Cyan
    if (-not (Test-IsVSCodeInstalled)) {
        if (Test-IsChocolateyInstalled) {
            Invoke-SubScript -SubScriptName "Install-VSCode.ps1" -StepDescription "Visual Studio Code Installation"
            if (-not (Test-IsVSCodeInstalled)) {
                 Write-Warning "VSCode not detected after installation attempt. This might be a PATH refresh delay for 'code.exe' or choco list timing."
            }
        } else {
            Write-Warning "Skipping Visual Studio Code installation: Chocolatey is not available."
            $Global:OverallSuccess = $false # VSCode is often a key dev tool
        }
    } else {
        Write-Host "Visual Studio Code already detected." -ForegroundColor Green
    }

    # 7. Python Packages 
    Write-Host "`n--- Processing Python Packages ---" -ForegroundColor Cyan
    if ($pythonSuccessfullyInstalledOrPresent -and (Test-IsPipInstalled)) {
        Invoke-SubScript -SubScriptName "Install-PythonPackages.ps1" -StepDescription "Python Packages Installation from requirements.txt"
    } else { 
        Write-Warning "Skipping Python packages installation."
        if (-not $pythonSuccessfullyInstalledOrPresent) { Write-Warning "Reason: Python is not available or not correctly configured." }
        elseif (-not (Test-IsPipInstalled)) { Write-Warning "Reason: Pip is not available." }
    }

    # 8. R Packages
    Write-Host "`n--- Processing R Packages ---" -ForegroundColor Cyan
    if ($Global:RScriptPath -and (Test-Path $Global:RScriptPath -PathType Leaf)) {
        Write-Host "Attempting to install R packages using Rscript at: $($Global:RScriptPath)"
        # Pass the RScriptPath to the Install-RPackages.ps1 script
        # The sub-script needs to be able to accept this as a parameter.
        Invoke-SubScript -SubScriptName "Install-RPackages.ps1" -StepDescription "R Packages Installation" 
    } elseif (Get-Command Rscript -ErrorAction SilentlyContinue) {
        Write-Host "Rscript found via Get-Command. Proceeding with R package installation."
        $Global:RScriptPath = (Get-Command Rscript).Source # Ensure it's set if found this way
        Invoke-SubScript -SubScriptName "Install-RPackages.ps1" -StepDescription "R Packages Installation"
    } else { 
        Write-Warning "Skipping R packages: Rscript.exe could not be located. Ensure R is installed and its 'bin' directory is on the PATH."
        Write-Warning "You may need to close this PowerShell window and open a new one for PATH changes to take effect, then re-run relevant parts or install R packages manually."
    }

    
    # 9. TinyTeX 
    Write-Host "`n--- Processing TinyTeX (for Quarto PDF output) ---" -ForegroundColor Cyan
    if (Test-IsQuartoCliInstalled) {
        if (-not (Test-IsTinyTeXInstalled)) {
            Invoke-SubScript -SubScriptName "Install-TinyTeX.ps1" -StepDescription "TinyTeX Installation (for Quarto)"
            if (-not (Test-IsTinyTeXInstalled)){
                 Write-Warning "TinyTeX not detected after installation attempt. PDF rendering via LaTeX might fail."
            }
        } else {
            Write-Host "TinyTeX already detected." -ForegroundColor Green
        }
    } else { 
        Write-Warning "Skipping TinyTeX: Quarto CLI is not available." 
    }
    
    

    # 10. 'nmfs-opensci/quarto_titlepages' Extension
    Write-Host "`n--- Processing nmfs-opensci/quarto_titlepages Extension ---" -ForegroundColor Cyan
    if (Test-IsQuartoCliInstalled) {
        if (-not (Test-IsNmfsExtensionInstalled -LocationOfExtensionParentDir $Global:nmfsExtensionInstallLocation)) {
            Write-Host "The 'nmfs-opensci/quarto_titlepages' extension will be installed in an '_extensions' folder within: $($Global:nmfsExtensionInstallLocation)"
            Invoke-SubScript -SubScriptName "Install-NmfsQuartoExtension.ps1" -StepDescription "'nmfs-opensci/quarto_titlepages' Quarto Extension Installation"
            if (-not (Test-IsNmfsExtensionInstalled -LocationOfExtensionParentDir $Global:nmfsExtensionInstallLocation)) {
                Write-Warning "nmfs-opensci/quarto_titlepages extension not detected after installation attempt."
            }
        } else {
            Write-Host "nmfs-opensci/quarto_titlepages extension already detected where script was run." -ForegroundColor Green
        }
    } else { 
        Write-Warning "Skipping 'nmfs-opensci/quarto_titlepages' extension: Quarto CLI is not available." 
    }

    # 11. Clone 'kamitor/quarto_titlepages' Repository 
    Write-Host "`n--- Processing kamitor/quarto_titlepages Repository ---" -ForegroundColor Cyan
    if (Test-IsGitInstalled) {
        if (-not (Test-IsKamitorRepoCloned)) {
            Invoke-SubScript -SubScriptName "Clone-KamitorRepo.ps1" -StepDescription "Cloning 'kamitor/quarto_titlepages' Repository"
            if (-not (Test-IsKamitorRepoCloned)) {
                 Write-Error "Failed to clone or detect 'kamitor/quarto_titlepages' repository."
                 $Global:OverallSuccess = $false # This repo is a primary goal
            }
        } else {
             Write-Host "'kamitor/quarto_titlepages' repository already cloned to default location." -ForegroundColor Green
        }
    } else { 
        Write-Warning "Skipping repository cloning: Git is not available."
        $Global:OverallSuccess = $false # Git is needed for this
    }
    
# This closes the main 'if ($Global:OverallSuccess -and (Test-IsChocolateyInstalled))' block from Corrected Part 3.
# The font and outlook checks will now run even if some choco installs failed,
# as long as the very initial Chocolatey setup itself was successful.
} elseif (-not (Test-IsChocolateyInstalled)) { 
    Write-Error "Cannot proceed with most tool installations because Chocolatey is not available."
    $Global:OverallSuccess = $false
}

 Write-Host "`n--- Processing RStudio IDE ---" -ForegroundColor Cyan
    if (Test-IsChocolateyInstalled) { # RStudio installation relies on Chocolatey
        if (-not (Test-IsRStudioInstalled)) {
            if (Invoke-SubScript -SubScriptName "Install-RStudio.ps1" -StepDescription "RStudio IDE Installation") {
                # No specific PATH refresh needed for RStudio GUI itself, but good to re-test
                Test-IsRStudioInstalled 
            } else {
                Write-Warning "RStudio installation sub-script reported an issue or was exited."
            }
        } else {
            Write-Host "RStudio already detected." -ForegroundColor Green
        }
    } else {
         Write-Warning "Skipping RStudio IDE installation: Chocolatey is not available."
         # Decide if this should set $Global:OverallSuccess = $false if RStudio is critical
    }

# 12. Install Custom Fonts
Write-Host "`n--- Processing Custom Fonts ---" -ForegroundColor Cyan
$fontDir = Join-Path -Path $Global:MainInstallerBaseDir -ChildPath "assets\fonts" 
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

# 13. Check for Microsoft Outlook
Write-Host "`n--- Checking for Microsoft Outlook ---" -ForegroundColor Cyan
if (-not (Test-IsOutlookInstalled)) {
    Write-Warning "Microsoft Outlook was not detected. If the project requires Outlook interaction, please ensure it is installed and configured manually."
} else {
    Write-Host "Microsoft Outlook appears to be installed. Please ensure it is configured with a mail profile if required by the project." -ForegroundColor Green
}

Write-Host "`n--- Installation Orchestration Attempted ---" -ForegroundColor Yellow
if ($Global:OverallSuccess) {
    Write-Host "All critical prerequisite steps appear to have completed successfully or were already satisfied." -ForegroundColor Green
    Write-Host "Please review any warnings above for non-critical items or PATH issues that might require a shell restart." -ForegroundColor Yellow
} else {
    Write-Error "One or more CRITICAL steps failed or critical prerequisites (like Python stub or Choco) were not resolved."
    Write-Warning "Please review the entire output above carefully. Some tools may have installed, but the environment is not fully set up."
    Write-Warning "Addressing CRITICAL errors (like Python stubs or failed core installs) and re-running, or restarting your shell, might be necessary."
}

Write-Host "`nIMPORTANT NOTES (Review Carefully):" -ForegroundColor Yellow
Write-Host " - PATH Environment Variable: Newly installed command-line tools (Python, Pip, R, Rscript, Git, Quarto)"
Write-Host "   might not be immediately available in THIS PowerShell session, even after refresh attempts."
Write-Host "   If you see 'command not found' errors when trying to use them after this script,"
Write-Host "   CLOSE THIS POWERSHELL WINDOW AND OPEN A NEW ADMINISTRATOR POWERSHELL WINDOW."
Write-Host "   This usually resolves PATH-related issues for new installations."
Write-Host " - Python App Execution Aliases (Stubs): If warnings about 'Windows Store stub' for Python appeared,"
Write-Host "   it is CRITICAL to manually disable these in Windows Settings ('Manage app execution aliases')"
Write-Host "   and then re-run this script in a new Admin PowerShell session for Python-dependent tools to work."
Write-Host " - Chocolatey Pending Reboot: If Chocolatey mentioned a pending reboot, some installations might not be fully"
Write-Host "   stable until the system is restarted, though the script attempts to continue."
Write-Host " - The 'nmfs-opensci/quarto_titlepages' extension was attempted to be installed into an '_extensions' folder"
Write-Host "   within the directory where this script was run: $($Global:nmfsExtensionInstallLocation)\_extensions"
Write-Host " - The main project files from '$ScriptScopeCloneRepoName' should be in: $($Global:FullClonePath) (if cloning was attempted)."
Write-Host "   Navigate there to use the project, e.g., 'cd ""$($Global:FullClonePath)""'."
Write-Host ""
Write-Host "Post-script actions that might still be needed depending on output:"
Write-Host " 1. Custom font installation was attempted. If fonts are not available, a system REBOOT or LOGOFF/LOGON might be necessary."
Write-Host " 2. Ensure Microsoft Outlook is configured with a mail profile if the project requires it."
Write-Host " 3. If Python/Pip or R/Rscript steps reported issues, manually verify their installation and PATH after restarting your shell."

Read-Host -Prompt "Script finished. Press Enter to exit."