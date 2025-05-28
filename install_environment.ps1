#Requires -Version 5.1

param()

$ProjectRoot = $PSScriptRoot
$PythonMinVersion = "3.8"
$RMinVersion = "4.0"
$PSMinVersion = "5.1"

$RequiredRPackages = @(
    "readr", "dplyr", "stringr", "tidyr", "ggplot2", "fmsb", "scales"
)

$RequiredPythonPackages = @(
    "pandas", "pywin32"
)

$CustomFontFileName = "QTDublinIrish.otf"
$CustomFontSourcePath = Join-Path $ProjectRoot "fonts" $CustomFontFileName
$SystemFontsDir = Join-Path $env:WINDIR "Fonts"
$InstalledFontPath = Join-Path $SystemFontsDir $CustomFontFileName

Function Write-Step { param($Message) Write-Host "`n--- $Message ---" -ForegroundColor Cyan }
Function Write-Success { param($Message) Write-Host "[SUCCESS] $Message" -ForegroundColor Green }
Function Write-Warning { param($Message) Write-Host "[WARNING] $Message" -ForegroundColor Yellow }
Function Write-ErrorMsg { param($Message) Write-Host "[ERROR] $Message" -ForegroundColor Red }
Function Write-Info { param($Message) Write-Host "[INFO] $Message" }
Function Write-Prompt { param($Message) Write-Host "[PROMPT] $Message" -ForegroundColor Magenta }

Function Test-CommandExists {
    param($CommandName)
    return [bool](Get-Command $CommandName -ErrorAction SilentlyContinue)
}

Function Get-CommandPath {
    param($CommandName)
    return (Get-Command $CommandName -ErrorAction SilentlyContinue).Source
}

Function Compare-SoftwareVersions {
    param(
        [string]$Version1,
        [string]$Version2
    )
    try {
        return [System.Version]$Version1 -ge [System.Version]$Version2
    } catch {
        Write-Warning "Could not compare versions '$Version1' and '$Version2'. Assuming OK."
        return $true
    }
}

Clear-Host
Write-Host "===================================================================" -ForegroundColor White
Write-Host "      Project Environment Setup for Resilience Reports      " -ForegroundColor White
Write-Host "===================================================================" -ForegroundColor White
Write-Host "This script will guide you through installing necessary software."
Write-Host "An internet connection is required for downloads."
Write-Host "Some steps may require Administrator privileges."
Write-Host "Script location: $ProjectRoot"
Write-Host "-------------------------------------------------------------------"

$GlobalAbort = $false
$GlobalAllGood = $true

Trap {
    Write-ErrorMsg "An unexpected error occurred: $($_.Exception.Message)"
    $GlobalAbort = $true
    $GlobalAllGood = $false
    Exit 1
}

$Continue = Read-Host "Press Enter to begin the setup, or Ctrl+C to cancel"
If ($GlobalAbort) { Exit 1 }

Write-Step "0. Verifying PowerShell Version"
If ($PSVersionTable.PSVersion -lt [System.Version]$PSMinVersion) {
    Write-ErrorMsg "Your PowerShell version is $($PSVersionTable.PSVersion). This script requires version $PSMinVersion or newer."
    Write-ErrorMsg "Please update PowerShell or run this script on a system with a compatible version."
    $GlobalAllGood = $false
    Read-Host "Press Enter to exit."
    Exit 1
} Else {
    Write-Success "PowerShell version $($PSVersionTable.PSVersion) is compatible."
}

$PythonExePath = $null
$RScriptExePath = $null
$QuartoExePath = $null

Function Install-SoftwareLoop {
    param(
        [string]$SoftwareName,
        [string]$CommandName,
        [string]$MinVersion,
        [string]$DownloadUrl,
        [string]$PathInstruction,
        [ref]$ExePathOut
    )
    $ExePath = Get-CommandPath $CommandName
    If ($ExePath) {
        Write-Info "$SoftwareName found at: $ExePath"
        try {
            $VersionOutput = ""
            If ($SoftwareName -eq "Python") { $VersionOutput = (python --version 2>&1).Trim() }
            If ($SoftwareName -eq "R") { $VersionOutput = (Rscript -e "cat(R.version.string)" 2>&1) }
            If ($SoftwareName -eq "Quarto") { $VersionOutput = (quarto --version 2>&1).Trim() }

            $CurrentVersion = ($VersionOutput -split ' ')[-1] # Basic parsing, might need adjustment
            If ($SoftwareName -eq "R") { $CurrentVersion = ($VersionOutput | Select-String -Pattern "version (\d+\.\d+\.\d+)" | ForEach-Object {$_.Matches.Groups[1].Value}) }


            If ($CurrentVersion) {
                 Write-Info "$SoftwareName version: $CurrentVersion"
                If ($MinVersion -and (-not (Compare-SoftwareVersions $CurrentVersion $MinVersion))) {
                    Write-Warning "Your $SoftwareName version $CurrentVersion is older than the recommended $MinVersion."
                    Write-Prompt "It's recommended to upgrade from: $DownloadUrl"
                    Read-Host "Press Enter to continue with the current version, or Ctrl+C to stop and upgrade."
                }
            } Else {
                Write-Warning "Could not determine $SoftwareName version automatically."
            }
            $ExePathOut.Value = $ExePath
            return $true
        } catch {
            Write-Warning "Found $SoftwareName, but could not verify version. Assuming it's functional."
            $ExePathOut.Value = $ExePath
            return $true
        }
    } Else {
        Write-Warning "$SoftwareName ('$CommandName') not found in your system PATH."
        Write-Prompt "Please download and install $SoftwareName $MinVersion or newer from: $DownloadUrl"
        Write-Prompt $PathInstruction
        Read-Host "After installing $SoftwareName (and ensuring it's in PATH), press Enter to re-check, or Ctrl+C to abort."
        $ExePath = Get-CommandPath $CommandName
        If ($ExePath) {
            Write-Success "$SoftwareName found after manual installation."
            $ExePathOut.Value = $ExePath
            return $true
        } Else {
            Write-ErrorMsg "$SoftwareName still not found. Please install it correctly and re-run this script."
            $GlobalAllGood = $false
            return $false
        }
    }
}

Write-Step "1. Checking Python Installation"
If (-not (Install-SoftwareLoop -SoftwareName "Python" -CommandName "python" -MinVersion $PythonMinVersion -DownloadUrl "https://www.python.org/downloads/windows/" -PathInstruction "IMPORTANT: During Python installation, ensure 'Add Python to PATH' is checked." -ExePathOut ([ref]$PythonExePath))) {
    $GlobalAbort = $true
}
If ($GlobalAbort) { Read-Host "Setup aborted. Press Enter to exit."; Exit 1 }


Write-Step "2. Checking/Installing Python Packages"
If ($PythonExePath) {
    $PipExe = Get-CommandPath "pip"
    If (-not $PipExe) { $PipExe = Get-CommandPath "pip3" }
    $PipCmd = if ($PipExe) { $PipExe } else { "$PythonExePath -m pip" }
    Write-Info "Using '$PipCmd' for Python package management."

    ForEach ($Package in $RequiredPythonPackages) {
        Write-Info "Checking Python package: $Package"
        $IsInstalledCheck = Invoke-Expression "$PipCmd show $Package" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        If ($LASTEXITCODE -eq 0) {
            Write-Success "Python package '$Package' is already installed."
        } Else {
            Write-Prompt "Python package '$Package' not found. Attempting to install..."
            Invoke-Expression "$PipCmd install $Package"
            If ($LASTEXITCODE -ne 0) {
                Write-ErrorMsg "Failed to install Python package '$Package'. Please check errors above and try manually: $PipCmd install $Package"
                $GlobalAllGood = $false
            } Else {
                $VerifyInstall = Invoke-Expression "$PipCmd show $Package" -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                If ($LASTEXITCODE -eq 0) {
                    Write-Success "Python package '$Package' installed and verified."
                } Else {
                    Write-ErrorMsg "Installed Python package '$Package', but verification failed. Please check manually."
                    $GlobalAllGood = $false
                }
            }
        }
    }
} Else { Write-Warning "Python not found, skipping Python package installation." }


Write-Step "3. Checking R Installation"
If (-not (Install-SoftwareLoop -SoftwareName "R" -CommandName "Rscript" -MinVersion $RMinVersion -DownloadUrl "https://cran.r-project.org/bin/windows/base/" -PathInstruction "Ensure R is added to PATH during installation or manually afterwards." -ExePathOut ([ref]$RScriptExePath))) {
    $GlobalAbort = $true
}
If ($GlobalAbort) { Read-Host "Setup aborted. Press Enter to exit."; Exit 1 }


Write-Step "4. Checking/Installing R Packages"
If ($RScriptExePath) {
    $RInstallScriptContent = @"
    options(Ncpus = max(1, parallel::detectCores(logical=FALSE) %/% 2))
    required_pkgs <- c('$(($RequiredRPackages -join "', '")')')
    installed_pkgs <- rownames(installed.packages())
    missing_pkgs <- required_pkgs[!required_pkgs %in% installed_pkgs]

    if (length(missing_pkgs) > 0) {
      cat("Attempting to install missing R packages:", paste(missing_pkgs, collapse=", "), "\n")
      tryCatch({
        install.packages(missing_pkgs, repos='https://cloud.r-project.org/')
        installed_pkgs_after <- rownames(installed.packages())
        still_missing <- missing_pkgs[!missing_pkgs %in% installed_pkgs_after]
        if (length(still_missing) > 0) {
          cat("ERROR: Failed to install some R packages:", paste(still_missing, collapse=", "), "\nPlease try installing them manually in R console: install.packages(c('", paste(still_missing, collapse="','"), "'))\n")
          quit(save="no", status=1)
        } else {
          cat("Successfully installed/verified all required R packages.\n")
          quit(save="no", status=0)
        }
      }, error = function(e) {
        cat("ERROR during R package installation:", conditionMessage(e), "\n")
        cat("Please try installing them manually in R console: install.packages(c('", paste(missing_pkgs, collapse="','"), "'))\n")
        quit(save="no", status=1)
      })
    } else {
      cat("All required R packages are already installed.\n")
      quit(save="no", status=0)
    }
"@
    $TempRScriptPath = Join-Path $env:TEMP "temp_install_r_packages.R"
    $RInstallScriptContent | Set-Content -Path $TempRScriptPath -Encoding UTF8 -Force
    Write-Info "Running R script to check/install packages. This may take some time..."
    & $RScriptExePath $TempRScriptPath
    If ($LASTEXITCODE -ne 0) {
        Write-ErrorMsg "R package installation encountered errors. Check R output above."
        $GlobalAllGood = $false
    } Else { Write-Success "R packages check/installation process completed." }
    Remove-Item $TempRScriptPath -ErrorAction SilentlyContinue -Force
} Else { Write-Warning "R (Rscript) not found, skipping R package installation." }


Write-Step "5. Checking Quarto Installation"
If (-not (Install-SoftwareLoop -SoftwareName "Quarto" -CommandName "quarto" -MinVersion "1.3" -DownloadUrl "https://quarto.org/docs/get-started/" -PathInstruction "The Quarto .msi installer for Windows should add it to PATH." -ExePathOut ([ref]$QuartoExePath))) {
    $GlobalAbort = $true
}
If ($GlobalAbort) { Read-Host "Setup aborted. Press Enter to exit."; Exit 1 }


Write-Step "6. Checking/Installing TinyTeX (LaTeX for Quarto)"
If ($QuartoExePath) {
    Write-Info "Quarto uses TinyTeX for PDF generation."
    Write-Info "Checking existing LaTeX setup with 'quarto check pdf'..."
    $QuartoCheckOutput = Invoke-Expression "$QuartoExePath check pdf" 2>&1
    Write-Host $QuartoCheckOutput -ForegroundColorGray

    If ($QuartoCheckOutput -match "LaTeX.+OK") {
        Write-Success "Quarto reports a working LaTeX installation for PDF generation."
    } Else {
        Write-Warning "Quarto check indicates LaTeX might not be fully configured or found."
        Write-Prompt "This step can take a significant amount of time (5-15 mins) and download ~200-300MB."
        $choice = Read-Host "Do you want to run 'quarto install tinytex' to install/reinstall LaTeX? (Y/N)"
        If ($choice -match '^[Yy]$') {
            Write-Host "Attempting to install/update TinyTeX via Quarto... This may require Administrator rights."
            try {
                Invoke-Expression "$QuartoExePath install tinytex"
                If ($LASTEXITCODE -ne 0) {
                    Write-ErrorMsg "TinyTeX installation via Quarto failed. Check output above."
                    Write-Warning "You might need to run this PowerShell script as Administrator, or install a full LaTeX distribution like MiKTeX manually."
                    $GlobalAllGood = $false
                } Else {
                    Write-Success "TinyTeX installation command executed by Quarto. Verifying..."
                    $QuartoCheckAfterOutput = Invoke-Expression "$QuartoExePath check pdf" 2>&1
                    Write-Host $QuartoCheckAfterOutput -ForegroundColorGray
                    If ($QuartoCheckAfterOutput -match "LaTeX.+OK") {
                         Write-Success "Quarto now reports a working LaTeX installation."
                    } Else {
                        Write-ErrorMsg "TinyTeX installation attempted, but 'quarto check pdf' still reports issues."
                        $GlobalAllGood = $false
                    }
                }
            } catch {
                Write-ErrorMsg "Error during 'quarto install tinytex': $($_.Exception.Message)"
                Write-Warning "You might need to run this PowerShell script as Administrator."
                $GlobalAllGood = $false
            }
        } Else { Write-Info "TinyTeX installation skipped by user. If PDF generation fails, install LaTeX manually or re-run this step." }
    }
} Else { Write-Warning "Quarto not found, cannot manage or check TinyTeX automatically." }


Write-Step "7. Custom Font Installation ($CustomFontFileName)"
If (-not (Test-Path $CustomFontSourcePath)) {
    Write-ErrorMsg "Custom font file '$CustomFontSourcePath' not found! Please ensure it's in the '$($ProjectRoot)\fonts' subfolder."
    $GlobalAllGood = $false
} Else {
    If (Test-Path $InstalledFontPath) {
        Write-Success "Font '$CustomFontFileName' appears to be already installed in system fonts."
    } Else {
        Write-Warning "Font '$CustomFontFileName' not found in system fonts."
        Write-Prompt "The project requires the font '$CustomFontFileName', located at: $CustomFontSourcePath"
        Write-Prompt "To install it:"
        Write-Prompt "  1. Open File Explorer and navigate to the '$($ProjectRoot)\fonts' folder."
        Write-Prompt "  2. Right-click on '$CustomFontFileName'."
        Write-Prompt "  3. Select 'Install' or 'Install for all users' (recommended, may need admin rights)."
        Read-Host "Press Enter once you have attempted to install the font."
        If (Test-Path $InstalledFontPath) {
            Write-Success "Font '$CustomFontFileName' now detected in system fonts."
        } Else {
            Write-Warning "Font '$CustomFontFileName' still not detected in system fonts after manual install attempt."
            Write-Warning "A system restart OR logging out and back in might be required for the font to be recognized."
            $GlobalAllGood = $false
        }
    }
}


Write-Step "8. Verifying Essential Project Files/Folders"
$EssentialFiles = @(
    "generate_reports.py", "send_emails.py", "example_3.qmd", "references.bib",
    (Join-Path "data" "cleaned_master.csv")
)
$EssentialDirs = @("data", "img", "tex", "fonts")
$ProjectStructureOK = $true

ForEach ($ItemName in $EssentialFiles) {
    $ItemPath = Join-Path $ProjectRoot $ItemName
    If (-not (Test-Path $ItemPath -PathType Leaf)) {
        Write-ErrorMsg "Essential project file not found: $ItemPath"
        $ProjectStructureOK = $false
        $GlobalAllGood = $false
    }
}
ForEach ($ItemName in $EssentialDirs) {
    $ItemPath = Join-Path $ProjectRoot $ItemName
    If (-not (Test-Path $ItemPath -PathType Container)) {
        Write-ErrorMsg "Essential project directory not found: $ItemPath"
        $ProjectStructureOK = $false
        $GlobalAllGood = $false
    }
}

If ($ProjectStructureOK) {
    Write-Success "Basic project file and folder structure seems OK."
} Else {
    Write-Warning "Some essential project files/folders are missing. Please ensure the project is complete."
}


Write-Step "9. Microsoft Outlook Prerequisite (for sending emails)"
Write-Info "The 'send_emails.py' script requires Microsoft Outlook to be installed and configured on this computer."
Write-Info "This script does not install Outlook. Please ensure it's set up if you intend to use the email sending feature."


Write-Host "`n===================================================================" -ForegroundColor White
Write-Host "                 Setup Check Complete!                 " -ForegroundColor White
Write-Host "===================================================================" -ForegroundColor White

If ($GlobalAllGood -and $PythonExePath -and $RScriptExePath -and $QuartoExePath) {
    Write-Success "Core software (Python, R, Quarto) and project structure appear to be configured."
    Write-Info "You should now be able to run the project scripts."
    Write-Info "`nTo generate reports:"
    Write-Info "  In a terminal (like PowerShell or Command Prompt) in this project folder ($ProjectRoot),"
    Write-Info "  run:  python generate_reports.py"
    Write-Info "`nTo send emails (after reports are generated and Outlook is configured):"
    Write-Info "  Run:  python send_emails.py"
    Write-Info ""
    Write-Warning "If PDF generation fails with font errors, ensure '$CustomFontFileName' was installed correctly (Step 7) and consider restarting your computer or logging out/in."
    Write-Warning "If LaTeX errors persist, ensure TinyTeX was installed successfully (Step 6) or consider a manual installation of MiKTeX/TeX Live."
} Else {
    Write-ErrorMsg "Some critical components are missing or setup encountered issues. Please review the messages above and address them."
    Write-ErrorMsg "Re-run this script after making corrections."
}

Read-Host "`nPress Enter to exit this setup script."