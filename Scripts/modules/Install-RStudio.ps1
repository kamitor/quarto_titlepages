# Install-RStudio.ps1
Write-Host "--- Attempting to Install RStudio using Chocolatey ---" -ForegroundColor Yellow

try {
    Write-Host "Executing: choco install rstudio -y --no-progress"
    choco install rstudio -y --no-progress
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Chocolatey command to install RStudio failed with exit code $LASTEXITCODE."
        Write-Warning "Please check Chocolatey logs for more details."
        # Optionally, you could set a global error flag or specific exit code here if main script needs to know
        exit 1 # Indicate failure
    }
    Write-Host "RStudio installation command executed via Chocolatey." -ForegroundColor Green
    Write-Host "If RStudio was installed, it should now be available."
    Write-Host "You might need to refresh your desktop or Start Menu to see the icon."
} catch {
    Write-Error "An error occurred during RStudio installation attempt: $($_.Exception.Message)"
    exit 1 # Indicate failure
}

# Check again after installation attempt
if (choco list --local-only --exact --name-only --limit-output rstudio -r | Select-String "rstudio") {
    Write-Host "RStudio installation confirmed by 'choco list'." -ForegroundColor Green
} else {
    Write-Warning "RStudio installation NOT confirmed by 'choco list' after install attempt."
}

exit 0 # Indicate success or that the command ran