# Script: modules/Install-CustomFonts.ps1
# Purpose: Installs custom fonts.
# REQUIRES: Administrator privileges (already ensured by Main-Installer.ps1)

Write-Host "Starting custom font installation..." -ForegroundColor Yellow

# Define fonts to install. Assumes font files are in a 'fonts' subdirectory
# relative to the Main-Installer.ps1 script.
$BaseDir = $Global:PSScriptRoot # PSScriptRoot of Main-Installer, if sub-script is dot-sourced
                              # If called with &, PSScriptRoot here is $ModulesDir
                              # Let's assume Main-Installer sets a global var for its own base dir
if (-not $Global:MainInstallerBaseDir) {
    Write-Error "Global variable MainInstallerBaseDir not set. Cannot determine font source path."
    exit 1
}
$FontSourceDir = Join-Path -Path $Global:MainInstallerBaseDir -ChildPath "assets\fonts" # Or just $Global:MainInstallerBaseDir if fonts are next to main script

$FontsToInstall = @(
    @{ Name = "QTDublinIrish"; FileName = "QTDublinIrish.otf" }
    # Add other fonts here if needed
)

$FontsDir = "$env:SystemRoot\Fonts"
$FontRegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
$UserFontRegPath = "HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts" # For user-specific registration

$overallSuccess = $true

foreach ($font in $FontsToInstall) {
    $fontFile = $font.FileName
    $fontName = $font.Name # This is usually the "Font Name" you see, not filename
    $sourceFontPath = Join-Path -Path $FontSourceDir -ChildPath $fontFile
    $destFontPath = Join-Path -Path $FontsDir -ChildPath $fontFile

    Write-Host "Processing font: $($font.Name) ($($font.FileName))"

    if (-not (Test-Path $sourceFontPath -PathType Leaf)) {
        Write-Warning "Font file not found at '$sourceFontPath'. Skipping."
        $overallSuccess = $false
        continue
    }

    # Check if font is already registered (simple check by registry value name)
    # A more robust check would involve enumerating font objects
    $regValueSystem = Get-ItemProperty -Path $FontRegPath -Name "$($fontName) (TrueType)" -ErrorAction SilentlyContinue
    $regValueUser = Get-ItemProperty -Path $UserFontRegPath -Name "$($fontName) (TrueType)" -ErrorAction SilentlyContinue
    
    if ($regValueSystem -or $regValueUser) {
        Write-Host "Font '$($fontName)' appears to be already registered. Verifying file..." -ForegroundColor Green
        if (Test-Path $destFontPath -PathType Leaf) {
             Write-Host "Font file '$destFontPath' also exists. Assuming installed." -ForegroundColor Green
             continue
        } else {
            Write-Warning "Font '$($fontName)' registered but file '$destFontPath' missing. Attempting reinstall."
        }
    }

    Write-Host "Installing font '$fontFile' to '$FontsDir'..."
    try {
        Copy-Item -Path $sourceFontPath -Destination $FontsDir -Force -ErrorAction Stop
        Write-Host "Font file copied." -ForegroundColor Green

        # Register the font in the HKLM registry for all users
        # The value name is typically "Font Name (TrueType/OpenType)" and value data is the filename.
        $registryValueName = "$($fontName) (OpenType)" # Adjust if it's TrueType, etc.
        If ($fontFile -like "*.ttf") { $registryValueName = "$($fontName) (TrueType)" }


        Write-Host "Registering font in HKLM: '$registryValueName' = '$fontFile'"
        Set-ItemProperty -Path $FontRegPath -Name $registryValueName -Value $fontFile -Type String -Force -ErrorAction Stop
        Write-Host "Font registered in HKLM." -ForegroundColor Green
        
        # Also register for current user to ensure immediate availability in some apps without logoff/reboot
        Write-Host "Registering font in HKCU: '$registryValueName' = '$fontFile'"
        if (-not (Test-Path $UserFontRegPath)) {
            New-Item -Path $UserFontRegPath -Force | Out-Null
        }
        Set-ItemProperty -Path $UserFontRegPath -Name $registryValueName -Value $fontFile -Type String -Force -ErrorAction Stop
        Write-Host "Font registered in HKCU." -ForegroundColor Green


        Write-Host "Font '$($font.Name)' installation attempted." -ForegroundColor Green
    } catch {
        Write-Error "Failed to install or register font '$($font.Name)': $($_.Exception.Message)"
        $overallSuccess = $false
    }
}

if (-not $overallSuccess) {
    Write-Warning "One or more fonts may not have been installed correctly."
    # Optional: Add a broadcast message for a system reboot/logoff if fonts aren't appearing immediately
    # Add-Type -AssemblyName System.Windows.Forms
    # [System.Windows.Forms.SystemInformation]::BroadcastCherylChangeEvent() # This is more for system settings
    # A proper way is often a reboot or logoff/logon for full font cache refresh across all apps.
    Write-Host "A reboot or logoff/logon may be required for all applications to see newly installed fonts." -ForegroundColor Yellow
    exit 1
}

Write-Host "Custom font installation process completed." -ForegroundColor Green
# Advise on reboot if needed
Write-Host "It is recommended to restart applications or log off/on for new fonts to be available everywhere." -ForegroundColor Yellow
exit 0