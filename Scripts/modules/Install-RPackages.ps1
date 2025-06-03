# Script: modules/Install-RPackages.ps1
# Purpose: Installs specified R packages using Rscript.

Write-Host "Starting R packages installation..." -ForegroundColor Yellow
Write-Host "This step can be very time-consuming, especially for packages like 'tidyverse'."

$rPackages = @(
    "tidyverse",
    "fmsb",
    "scales",
    "rmarkdown",
    "knitr",
    "openxlsx",
    "readxl"
)

# Check if Rscript is available
if (-not (Get-Command Rscript -ErrorAction SilentlyContinue)) {
    Write-Error "Rscript command not found. Cannot install R packages."
    exit 1
}

Write-Host "Found Rscript: $(Rscript --version 2>&1 | Out-String). Attempting to install: $($rPackages -join ', ')"

# Construct the R command
# Using max(1, parallel::detectCores(logical=FALSE)-1) to avoid issues if detectCores returns 1 or errors.
# Setting Ncpus globally via options() before install.packages()
$rCommand = "options(Ncpus = tryCatch({max(1, parallel::detectCores(logical=FALSE)-1)}, error=function(e){1})); install.packages(c('" + ($rPackages -join "','") + "'), repos='http://cran.rstudio.com/', quiet=FALSE, verbose=TRUE)"

Write-Host "Executing R command: $rCommand" -ForegroundColor DarkGray

try {
    Rscript -e $rCommand
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "R packages installation (Rscript) failed with exit code $LASTEXITCODE."
        Write-Warning "Some packages might not have been installed. Check Rscript output above."
        exit 1
    }

    Write-Host "R packages installation command executed. Please check output for any individual package errors." -ForegroundColor Green
    exit 0
} catch {
    Write-Error "An unexpected error occurred during R packages installation: $($_.Exception.Message)"
    exit 1
}