Clear-Host
$Art = @"
  _        _                                  ____        _          _____ _ _            _             
 | |      | |                                / ___|  ___ | | __ _   / ____| (_)          | |            
 | |      | | __ _ _ __ ___  _ __ ___   ___  | |     / _ \| |/ _` | | |    | |_ _ __   ___| | ___   __ _ 
 | |  _   | |/ _` | '_ ` _ \| '_ ` _ \ / _ \ | |___ | (_) | | (_| | | |    | | | '_ \ / _ \ |/ _ \ / _` |
 | |_| |__| | (_| | | | | | | | | | | |  __/  \____| \___/|_|\__,_|  \_ \__|_|_| .__/ \___/_|\_ _|\__\_|
  \_____/\_/\__,_|_| |_| |_|_| |_| |_|\___|                           \_____| | |                      
                                                                             |_|                      
                        L E C T O R A A T   S U P P L Y   C H A I N   F I N A N C E
"@
Write-Host $Art -ForegroundColor Green
Write-Host ("-" * 80)
Write-Host "Environment Setup for kamitor/quarto_titlepages, initiated by Lectoraat Supply Chain Finance" -ForegroundColor Yellow
Write-Host "This script will attempt to install necessary tools and configure your system."
Write-Host ("-" * 80)
Write-Host " V 0.8"

Write-Host "--- Performing Pre-flight System Checks ---" -ForegroundColor Magenta

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

Write-Host "Checking Operating System version..." -ForegroundColor Gray
$OS = Get-CimInstance Win32_OperatingSystem
Write-Host "Operating System: $($OS.Caption), Version: $($OS.Version)"
if (($OS.Version -split "\.")[0] -lt 10) {
    Write-Warning "This script is optimized for Windows 10 or newer. You are on an older version ($($OS.Caption)). Some features might not work as expected."
} else {
    Write-Host "OS version check passed (Windows 10 or newer)." -ForegroundColor Green
}

Write-Host "Checking PowerShell version..." -ForegroundColor Gray
Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)"
if ($PSVersionTable.PSVersion.Major -lt 5 -or ($PSVersionTable.PSVersion.Major -eq 5 -and $PSVersionTable.PSVersion.Minor -lt 1) ) {
    Write-Warning "PowerShell 5.1 or higher is recommended. Your version is $($PSVersionTable.PSVersion)."
    Write-Warning "Consider upgrading PowerShell for best compatibility: https://aka.ms/PSWindows"
} else {
    Write-Host "PowerShell version check passed (5.1+)." -ForegroundColor Green
}

Write-Host "Checking for active Internet connection (pinging google.com)..." -ForegroundColor Gray
if (Test-Connection -ComputerName "google.com" -Count 1 -Quiet -ErrorAction SilentlyContinue) {
    Write-Host "Internet connection appears to be active." -ForegroundColor Green
} else {
    Write-Warning "Could not confirm an active Internet connection by pinging google.com."
    Write-Warning "An internet connection is required to download tools and packages."
    $proceedWithoutInternet = Read-Host "Do you want to attempt to continue anyway? (yes/no)"
    if ($proceedWithoutInternet -ne 'yes') {
        Write-Error "Exiting script as per user request due to unconfirmed internet connection."
        exit 1
    }
    Write-Warning "Proceeding without confirmed internet connection at user's risk."
}
Write-Host ("-" * 80)
Write-Host ""
Read-Host "Pre-flight checks complete. Press Enter to begin the installation process..."
Write-Host ""

try {
    Set-ExecutionPolicy Bypass -Scope Process -Force -ErrorAction Stop
    Write-Host "Execution policy set to Bypass for the current process." -ForegroundColor Green
} catch {
    Write-Error "Failed to set execution policy. Error: $($_.Exception.Message)"
    Read-Host "Press Enter to exit"
    exit 1
}

$BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$Global:MainInstallerBaseDir = $BaseDir
$ModulesDir = Join-Path -Path $BaseDir -ChildPath "modules"

$Global:OverallSuccess = $true
$Global:nmfsExtensionInstallLocation = $BaseDir

$ScriptScopeCloneRepoName = "kamitor_quarto_titlepages"
$Global:CloneParentDir = Join-Path -Path $HOME -ChildPath "Documents\GitHub"
$Global:FullClonePath = Join-Path -Path $Global:CloneParentDir -ChildPath $ScriptScopeCloneRepoName
$Global:RepoUrl = "https://github.com/kamitor/quarto_titlepages.git"

Function Invoke-SubScript {
    param(
        [string]$SubScriptName,
        [string]$StepDescription
    )
    Write-Host "`n--- Checking/Initiating: $StepDescription ---" -ForegroundColor Cyan
    $SubScriptPath = Join-Path -Path $ModulesDir -ChildPath $SubScriptName
    if (-not (Test-Path $SubScriptPath -PathType Leaf)) {
        Write-Error "Sub-script not found or is not a file: $SubScriptPath"
        $Global:OverallSuccess = $false
        return $false
    }

    try {
        Write-Host "Executing sub-script: $SubScriptPath"
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
    $chocoProfilePath = Join-Path -Path $env:ProgramData -ChildPath "chocolatey\helpers\chocolateyProfile.psm1"
    if (Test-Path $chocoProfilePath) {
        Import-Module $chocoProfilePath -ErrorAction SilentlyContinue
    }
    Write-Host "PATH refresh attempted." -ForegroundColor DarkGray
}

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
            $vscodePackageInfo = choco list --local-only --exact --name-only --limit-output vscode -r 
        } catch {
             Write-Warning "Error checking for VSCode with choco list: $($_.Exception.Message)"
        }
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
    if (Test-Path (Join-Path $tinytexPathUser "bin" "win32" "pdflatex.exe") -PathType Leaf) {
        Write-Host "TinyTeX found at user path: $tinytexPathUser" -ForegroundColor Green
        return $true
    }
    
    Write-Host "TinyTeX not found at common user path. Trying 'quarto check' for LaTeX detection..."
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
        Write-Warning "Exception during Quarto check: $($_.Exception.Message)"
        $exitCode = -1 
    } finally {
        $ProgressPreference = $OriginalProgressPreference
    }

    if ($exitCode -ne 0) { 
        Write-Host "Attempt to use 'quarto check' for TinyTeX detection failed or did not find LaTeX. Last ExitCode: $exitCode" -ForegroundColor Yellow
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
        (Join-Path $env:LOCALAPPDATA "Microsoft\WindowsApps\outlookforwindows.exe"),
        (Join-Path $env:ProgramW6432 "Microsoft Office\root\Office16\OUTLOOK.EXE") 
    )
    $uniqueOutlookPaths = $outlookPaths | Select-Object -Unique

    $outlookAppxPackage = Get-AppxPackage -Name "microsoft.outlookforwindows" -ErrorAction SilentlyContinue
            
    $foundPath = $null
    foreach ($path in $uniqueOutlookPaths) {
        if ($path -and (Test-Path $path -PathType Leaf)) { 
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

Write-Host "`nStarting installation checks and process..."

if (-not (Test-IsChocolateyInstalled)) {
    if (-not (Invoke-SubScript -SubScriptName "Install-Chocolatey.ps1" -StepDescription "Chocolatey Package Manager Installation")) {
    }
    Refresh-CurrentSessionPath
    if (-not (Test-IsChocolateyInstalled)) {
        Write-Error "FATAL: Chocolatey installation failed or was not detected after install attempt. Cannot proceed."
        $Global:OverallSuccess = $false
    }
}

if ($Global:OverallSuccess -and (Test-IsChocolateyInstalled)) {
    Write-Host "Chocolatey detected. Ensuring Chocolatey executables are findable in this session." -ForegroundColor DarkGray
    $chocoPath = Join-Path -Path $env:ProgramData -ChildPath "chocolatey\bin"
    if (Test-Path $chocoPath -and ($env:Path -notlike "*$chocoPath*")) {
        Write-Host "Adding Chocolatey bin to session PATH: $chocoPath"
        $env:Path = "$chocoPath;$($env:Path)"
    }
    $chocoProfilePathGlobal = Join-Path -Path $env:ProgramData -ChildPath "chocolatey\helpers\chocolateyProfile.psm1"
    if (Test-Path $chocoProfilePathGlobal) {
        Import-Module $chocoProfilePathGlobal -Force -ErrorAction SilentlyContinue
    }
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        Write-Warning "Choco command still not found by Get-Command after explicit PATH manipulation. Sub-scripts using 'choco' might fail."
    }
}

if ($Global:OverallSuccess -and (Test-IsChocolateyInstalled)) {

    if (-not (Test-IsGitInstalled)) {
        if (Invoke-SubScript -SubScriptName "Install-Git.ps1" -StepDescription "Git Installation") {
            Refresh-CurrentSessionPath
        } 
        Test-IsGitInstalled 
    }
}

Function Test-IsTinyTeXInstalled {
    Write-Host "Checking for TinyTeX..."
    if (-not (Test-IsQuartoCliInstalled)) {
        Write-Host "Quarto CLI not installed, cannot check for TinyTeX." -ForegroundColor Yellow
        return $false
    }

    $tinytexPathUser = Join-Path $env:APPDATA "TinyTeX"
    if (Test-Path (Join-Path $tinytexPathUser "bin" "win32" "pdflatex.exe") -PathType Leaf) {
        Write-Host "TinyTeX found at user path: $tinytexPathUser" -ForegroundColor Green
        return $true
    }
    
    Write-Host "TinyTeX not found at common user path. Trying 'quarto check' for LaTeX detection..."
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
        Write-Warning "Exception during Quarto check: $($_.Exception.Message)"
        $exitCode = -1 
    } finally {
        $ProgressPreference = $OriginalProgressPreference
    }

    if ($exitCode -ne 0) { 
        Write-Host "Attempt to use 'quarto check' for TinyTeX detection failed or did not find LaTeX. Last ExitCode: $exitCode" -ForegroundColor Yellow
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
        (Join-Path $env:LOCALAPPDATA "Microsoft\WindowsApps\outlookforwindows.exe"),
        (Join-Path $env:ProgramW6432 "Microsoft Office\root\Office16\OUTLOOK.EXE") 
    )
    $uniqueOutlookPaths = $outlookPaths | Select-Object -Unique

    $outlookAppxPackage = Get-AppxPackage -Name "microsoft.outlookforwindows" -ErrorAction SilentlyContinue
            
    $foundPath = $null
    foreach ($path in $uniqueOutlookPaths) {
        if ($path -and (Test-Path $path -PathType Leaf)) { 
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

Write-Host "`nStarting installation checks and process..."

if (-not (Test-IsChocolateyInstalled)) {
    if (-not (Invoke-SubScript -SubScriptName "Install-Chocolatey.ps1" -StepDescription "Chocolatey Package Manager Installation")) {
    }
    Refresh-CurrentSessionPath
    if (-not (Test-IsChocolateyInstalled)) {
        Write-Error "FATAL: Chocolatey installation failed or was not detected after install attempt. Cannot proceed."
        $Global:OverallSuccess = $false
    }
}

if ($Global:OverallSuccess -and (Test-IsChocolateyInstalled)) {
    Write-Host "Chocolatey detected. Ensuring Chocolatey executables are findable in this session." -ForegroundColor DarkGray
    $chocoPath = Join-Path -Path $env:ProgramData -ChildPath "chocolatey\bin"
    if (Test-Path $chocoPath -and ($env:Path -notlike "*$chocoPath*")) {
        Write-Host "Adding Chocolatey bin to session PATH: $chocoPath"
        $env:Path = "$chocoPath;$($env:Path)"
    }
    $chocoProfilePathGlobal = Join-Path -Path $env:ProgramData -ChildPath "chocolatey\helpers\chocolateyProfile.psm1"
    if (Test-Path $chocoProfilePathGlobal) {
        Import-Module $chocoProfilePathGlobal -Force -ErrorAction SilentlyContinue
    }
    if (-not (Get-Command choco -ErrorAction SilentlyContinue)) {
        Write-Warning "Choco command still not found by Get-Command after explicit PATH manipulation. Sub-scripts using 'choco' might fail."
    }
}

if ($Global:OverallSuccess -and (Test-IsChocolateyInstalled)) {

    if (-not (Test-IsGitInstalled)) {
        if (Invoke-SubScript -SubScriptName "Install-Git.ps1" -StepDescription "Git Installation") {
            Refresh-CurrentSessionPath
        } 
        Test-IsGitInstalled 
    }