# Script: modules/Install-PythonPackages.ps1
# Purpose: Installs specified Python packages using pip.

Write-Host "Starting Python packages installation..." -ForegroundColor Yellow

$pythonPackages = @(
    "pandas",
    "pywin32",
    "numpy",
    "scipy",
    "matplotlib",
    "seaborn",
    "openpyxl",
    "requests"
)

# Check if pip is available
if (-not (Get-Command pip -ErrorAction SilentlyContinue)) {
    Write-Error "pip command not found. Cannot install Python packages."
    exit 1
}

Write-Host "Found pip: $(pip --version)"
Write-Host "Attempting to install the following Python packages: $($pythonPackages -join ', ')"

try {
    # Upgrade pip first
    Write-Host "Attempting to upgrade pip..."
    python.exe -m pip install --upgrade pip --user
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "Failed to upgrade pip. Continuing with existing version. Exit code: $LASTEXITCODE"
    } else {
        Write-Host "pip upgrade attempt finished." -ForegroundColor Green
    }

    # Install packages
    Write-Host "Executing: pip install $($pythonPackages -join ' ')"
    pip install $pythonPackages
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "pip install command failed with exit code $LASTEXITCODE."
        # List packages that might have failed
        Write-Warning "Some packages might not have been installed. Check pip output above."
        exit 1
    }

    Write-Host "Python packages installation command executed successfully." -ForegroundColor Green
    exit 0
} catch {
    Write-Error "An unexpected error occurred during Python packages installation: $($_.Exception.Message)"
    exit 1
}