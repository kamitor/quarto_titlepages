# Main-Installer.ps1
# Orchestrates the installation of all necessary tools and programs.
# REQUIRES: Run this script in PowerShell as Administrator.
# NOTE: VSCode uses PSScriptAnalyzer for linting. Some "errors" it reports might be style/best-practice warnings.
# If a true syntax error occurs, PowerShell itself will likely fail to parse/run the script.

# --- Initial Setup ---
Write-Host "--- Starting Main Installation Orchestrator ---" -ForegroundColor Yellow

# Ensure running as Administrator
$currentUser = New-Object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())
if (-Not $currentUser.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "This script must be run as Administrator. Please re-launch PowerShell as Administrator and try again."
    Read-Host "Press Enter to exit"
    exit 1
}
Write-Host "Administrator privileges confirmed." -ForegroundColor Green

# Set Execution Policy for this process to allow script execution
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
        # Try to get specific info for vscode package
        $vscodePackageInfo = choco list --local-only --exact --name-only --limit-output vscode
        if ($LASTEXITCODE -eq 0 -and $vscodePackageInfo -match "vscode") {
            Write-Host "VSCode (Chocolatey package 'vscode') is listed as installed by Chocolatey." -ForegroundColor Green
            # Additionally, check if the executable is on path, as a secondary confirmation
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

    $tinytexPathUser = Join-Path $env:APPDATA "TinyTeX" # Common path for 'quarto install tool tinytex'
    $tinytexPathQuartoBundled = "" # Placeholder for system-wide if Quarto ever bundles it differently

    # First, check common installation paths
    if (Test-Path (Join-Path $tinytexPathUser "bin" "win32" "pdflatex.exe") -PathType Leaf) {
        Write-Host "TinyTeX found at user path: $tinytexPathUser" -ForegroundColor Green
        return $true
    }
    
    # Add other known paths if necessary, e.g.:
    # $quartoInstallPath = (Get-Command quarto).Source | Split-Path | Split-Path
    # $tinytexPathQuartoBundled = Join-Path $quartoInstallPath "share\tinytex"
    # if (Test-Path (Join-Path $tinytexPathQuartoBundled "bin" "win32" "pdflatex.exe") -PathType Leaf) {
    #     Write-Host "TinyTeX found bundled with Quarto at: $tinytexPathQuartoBundled" -ForegroundColor Green
    #     return $true
    # }

    # As a fallback, try 'quarto check' and parse its output (can be version dependent)
    Write-Host "TinyTeX not found at common paths. Trying 'quarto check' for LaTeX detection..."
    $quartoCheckOutput = ""
    $exitCode = 1 # Default to failure
    try {
        $OriginalProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        # Use a more general check that should list LaTeX status
        $quartoCheckOutput = quarto check --json 2>$null # Try to get JSON output, easier to parse
        if ($LASTEXITCODE -eq 0 -and $quartoCheckOutput) {
             $checkResult = $quartoCheckOutput | ConvertFrom-Json
             if ($checkResult.formats.pdf.latex) {
                Write-Host "Quarto check indicates a LaTeX distribution is available: $($checkResult.formats.pdf.latex)" -ForegroundColor Green
                # This doesn't guarantee it's TinyTeX, but it's a functional LaTeX
                return $true
             } else {
                Write-Host "Quarto check JSON does not confirm LaTeX." -ForegroundColor Yellow
             }
        } else {
            # Fallback to string parsing if JSON failed or no output
            $quartoCheckOutput = quarto check 2>&1 | Out-String
            if ($LASTEXITCODE -eq 0 -and $quartoCheckOutput -match "(?i)LaTeX\s+\[✔\]|(?i)Found LaTeX") {
                 Write-Host "Quarto check (text) indicates LaTeX is available." -ForegroundColor Green
                 return $true
            } else {
                Write-Host "Quarto check (text) does not confirm LaTeX. Output: $($quartoCheckOutput | Select-Object -First 300)" -ForegroundColor Yellow
            }
        }
        $exitCode = $LASTEXITCODE
    } catch {
        $quartoCheckOutput = $_.Exception.Message
        $exitCode = -1
    } finally {
        $ProgressPreference = $OriginalProgressPreference
    }

    if ($exitCode -ne 0) {
        Write-Host "Attempt to use 'quarto check' for TinyTeX detection failed or did not find LaTeX. ExitCode: $exitCode" -ForegroundColor Yellow
    }
    
    Write-Host "TinyTeX not definitively detected." -ForegroundColor Yellow
    return $false
}

Function Test-IsNmfsExtensionInstalled {
    param (
        [string]$LocationOfExtensionParentDir # This is $Global:nmfsExtensionInstallLocation
    )
    $extensionDirName = "nmfs-opensci" # Top-level directory Quarto creates for this extension
    $fullExtensionPath = Join-Path -Path $LocationOfExtensionParentDir -ChildPath "_extensions\$extensionDirName"
    Write-Host "Checking for '$extensionDirName' Quarto extension in '$fullExtensionPath'..."

    if (Test-Path $fullExtensionPath -PathType Container) { # Check if it's a directory
        Write-Host "'$extensionDirName' extension found at $fullExtensionPath." -ForegroundColor Green
        return $true
    }
    Write-Host "'$extensionDirName' extension NOT found in $fullExtensionPath." -ForegroundColor Yellow
    return $false
}

Function Test-IsKamitorRepoCloned {
    # Uses $ScriptScopeCloneRepoName (from script scope) and $Global:FullClonePath (from global scope)
    Write-Host "Checking if '$ScriptScopeCloneRepoName' repository is cloned to '$($Global:FullClonePath)'..."
    if (Test-Path (Join-Path -Path $Global:FullClonePath -ChildPath ".git") -PathType Container) { # Check for .git folder
        Write-Host "Repository found at $($Global:FullClonePath)." -ForegroundColor Green
        return $true
    }
    Write-Host "Repository NOT found at $($Global:FullClonePath) (or is not a git repository)." -ForegroundColor Yellow
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
# Ensure Chocolatey profile is loaded if already installed or just installed
$chocoProfilePathGlobal = Join-Path -Path $env:ProgramData -ChildPath "chocolatey\helpers\chocolateyProfile.psm1"
if ($Global:OverallSuccess -and (Test-Path $chocoProfilePathGlobal)) { # Only if previous steps allow
    Import-Module $chocoProfilePathGlobal -ErrorAction SilentlyContinue
}


# Proceed only if Chocolatey is available and previous steps were successful
if ($Global:OverallSuccess -and (Test-IsChocolateyInstalled)) {

    # 2. Git
    if (-not (Test-IsGitInstalled)) {
        if (Invoke-SubScript -SubScriptName "Install-Git.ps1" -StepDescription "Git Installation") {
            Refresh-CurrentSessionPath
        } # Error handled by Invoke-SubScript
        Test-IsGitInstalled # Re-test and display status
    }

    # 3. Python & Pip
    $pythonSuccessfullyInstalledOrPresent = Test-IsPythonInstalled
    if (-not $pythonSuccessfullyInstalledOrPresent) {
        Write-Host "Python not found. Attempting installation via sub-script."
        if (Invoke-SubScript -SubScriptName "Install-Python.ps1" -StepDescription "Python Installation") {
            Refresh-CurrentSessionPath
            if (Test-IsPythonInstalled) {
                $pythonSuccessfullyInstalledOrPresent = $true
                Write-Host "Python is now detected after installation attempt." -ForegroundColor Green
            } else {
                Write-Error "Python installation was attempted but Python is still not detected."
                $Global:OverallSuccess = $false # Explicitly mark failure for this critical step
            }
        } else {
             $Global:OverallSuccess = $false # Python install sub-script failed
        }
    }

    if ($pythonSuccessfullyInstalledOrPresent -and $Global:OverallSuccess) {
        if (-not (Test-IsPipInstalled)) {
            Write-Warning "Python is installed, but Pip was not found. Attempting to install/ensure Pip using 'python -m ensurepip'..."
            try {
                Write-Host "Executing: python -m ensurepip --upgrade" -ForegroundColor DarkGray
                & python -m ensurepip --upgrade
                if ($LASTEXITCODE -ne 0) {
                    Write-Error "Attempt to run 'python -m ensurepip --upgrade' failed with exit code $LASTEXITCODE."
                    # $Global:OverallSuccess = $false # Consider if this is fatal for the whole script
                } else {
                    Write-Host "'python -m ensurepip --upgrade' executed. Refreshing PATH and re-checking for Pip." -ForegroundColor Green
                    Refresh-CurrentSessionPath
                    Test-IsPipInstalled # Re-check and display status
                }
            } catch {
                 Write-Error "An error occurred while trying to run 'python -m ensurepip --upgrade': $($_.Exception.Message)"
                 # $Global:OverallSuccess = $false
            }
        }
    }

    # 4. R
    if ($Global:OverallSuccess -and -not (Test-IsRInstalled)) {
        if (Invoke-SubScript -SubScriptName "Install-R.ps1" -StepDescription "R Installation") {
            Refresh-CurrentSessionPath
        }
        Test-IsRInstalled
    }

    # 5. Quarto CLI
    if ($Global:OverallSuccess -and -not (Test-IsQuartoCliInstalled)) {
        if (Invoke-SubScript -SubScriptName "Install-QuartoCli.ps1" -StepDescription "Quarto CLI Installation") {
            Refresh-CurrentSessionPath
        }
        Test-IsQuartoCliInstalled
    }

    # 6. Visual Studio Code
    if ($Global:OverallSuccess -and -not (Test-IsVSCodeInstalled)) {
        Invoke-SubScript -SubScriptName "Install-VSCode.ps1" -StepDescription "Visual Studio Code Installation"
        Test-IsVSCodeInstalled # Re-test, path refresh not usually needed for VSCode app itself
    }

    # --- Language Specific Packages & Tools (after main tools are confirmed) ---

    # 7. Python Packages
    if ($Global:OverallSuccess -and $pythonSuccessfullyInstalledOrPresent -and (Test-IsPipInstalled)) {
        Invoke-SubScript -SubScriptName "Install-PythonPackages.ps1" -StepDescription "Python Packages Installation"
    } elseif ($Global:OverallSuccess) { # Only show warning if we haven't failed out earlier
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

} elseif (-not (Test-IsChocolateyInstalled)) { # If the initial Choco check failed and install also failed.
    Write-Error "Cannot proceed with tool installations because Chocolatey is not available."
    # $Global:OverallSuccess is already false from the Choco install check
}

# 12. Install Custom Fonts
# Create an 'assets/fonts' subdirectory in the same location as Main-Installer.ps1
# and place QTDublinIrish.otf (and any other required fonts) there.
$fontDir = Join-Path -Path $Global:MainInstallerBaseDir -ChildPath "assets\fonts"
if (Test-Path $fontDir -PathType Container) {
    if (Get-ChildItem -Path $fontDir -Filter "*.otf" -ErrorAction SilentlyContinue) { # Or *.ttf etc.
        Invoke-SubScript -SubScriptName "Install-CustomFonts.ps1" -StepDescription "Custom Font Installation"
    } else {
        Write-Host "No font files found in '$fontDir'. Skipping custom font installation." -ForegroundColor DarkGray
    }
} else {
    Write-Warning "Font directory '$fontDir' not found. Skipping custom font installation."
    Write-Warning "Create the directory and place font files (e.g., QTDublinIrish.otf) there if needed."
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
Write-Host "Manual steps still potentially required:"
Write-Host " 1. Install custom font (e.g., QTDublinIrish.otf) if needed by the project."
Write-Host " 2. Ensure Microsoft Outlook is configured if the project requires it."

Read-Host -Prompt "Script finished. Press Enter to exit."