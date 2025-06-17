param (
    [string]$ExplicitRScriptPath = $Global:RScriptPath 
)

Write-Host "--- Starting R Package Installation (using R.exe CMD R) ---" -ForegroundColor Yellow

$rPackages = @(
    # --- Core Quarto/R Markdown & Reporting Essentials ---
    "rmarkdown", 
    "knitr",    
    "yaml",         
    "jsonlite",     
    "htmltools",    
    "evaluate",     
    "xfun",         
    "DT",               # For interactive HTML tables
    "kableExtra",       # For creating beautiful static tables 
    "pagedown",         # For creating paged HTML documents (e.g., resumes, letters)
    "bookdown",         # For writing books and long-form documents
    "blogdown",         # For creating websites with R Markdown
    "distill",          # For scientific and technical writing websites
    "shiny",            # For interactive web applications
    "flexdashboard",    # For creating dashboards

    # --- Tidyverse Ecosystem (Data Science Core) ---
    "tidyverse",        # Meta-package for dplyr, ggplot2, tidyr, readr, purrr, tibble, stringr, forcats
    "readr",            # (dependency of tidyverse) For reading rectangular data
    "dplyr",            # (dependency of tidyverse) Data manipulation
    "stringr",          # (dependency of tidyverse) String manipulation
    "tidyr",            # (dependency of tidyverse) Tidying data
    "ggplot2",          # (dependency of tidyverse) Advanced plotting
    "purrr",            # (dependency of tidyverse) Functional programming
    "tibble",           # (dependency of tidyverse) Modern data frames
    "forcats",          # (dependency of tidyverse) Tools for factors
    
    # --- Data Import/Export & Management ---
    "readxl",           # For reading Excel files (.xls, .xlsx)
    "writexl",          # For writing to Excel files
    "openxlsx",         # More advanced Excel reading/writing/formatting
    "haven",            # For reading SPSS, Stata, and SAS files
    "DBI",              # Database Interface (common API for database connections)
    "RSQLite",          # SQLite database driver
    "RPostgres",        # PostgreSQL database driver (example, add others like RMySQL, odbc as needed)
    "feather",          # Fast, lightweight binary format for data frames
    "fst",              # Another fast data frame storage format
    "arrow",            # For working with Apache Arrow format, efficient for large data & interoperability

    # --- Data Visualization (beyond base ggplot2) ---
    "ggrepel",          # For preventing text label overlap in ggplot2
    "patchwork",        # For combining ggplot2 plots easily
    "plotly",           # For interactive D3.js-based plots
    "highcharter",      # Another library for interactive charts
    "leaflet",          # For interactive maps
    "sf",               # For working with spatial vector data (Simple Features)
    "tmap",             # For thematic maps
    "gganimate",        # For creating animations with ggplot2
    "viridis",          # Colorblind-friendly color palettes
    "RColorBrewer",     # Palettes for cartography and data visualization
    "ggthemes",         # Additional themes for ggplot2
    "cowplot",          # Helpers for ggplot2, arranging plots

    # --- Modeling & Machine Learning (Selection) ---
    "caret",            # Classification And REgression Training (meta-package for ML)
    "tidymodels",       # Newer meta-package for modeling, follows tidyverse principles
    "ranger",           # Fast Random Forest implementation
    "xgboost",          # Extreme Gradient Boosting
    "glmnet",           # LASSO and Elastic-Net Regularized Generalized Linear Models
    "randomForest",     # Classic Random Forest
    "rpart",            # Recursive Partitioning and Regression Trees
    "prophet",          # Time series forecasting by Facebook
    "forecast",         # Comprehensive time series forecasting tools
    "tsibble",          # Tidy data structures for time series
    "fable",            # Tidy time series forecasting models

    # --- Text Analysis & NLP (Selection) ---
    "tidytext",         # Text mining using tidy data principles
    "tm",               # Text Mining package (framework)
    "quanteda",         # Quantitative analysis of textual data
    "SnowballC",        # Stemming C library
    
    # --- Utility, Time, and Other Useful Packages ---
    "lubridate",        # For dates and times
    "data.table",       # Alternative for fast data manipulation
    "fmsb",             # Already in your list (for radar charts)
    "scales",           # Already in your list (often used with ggplot2 for scaling)
    "summarytools",     # For quick and comprehensive summary statistics
    "janitor",          # For simple data cleaning
    "skimr",            # Compact and flexible summaries of data
    "here",             # For constructing paths relative to project root, makes projects portable
    "fs",               # File system operations
    "httr",             # For working with HTTP requests (if you fetch data from APIs)
    "rvest",            # For web scraping
    "xml2",             # For working with XML
    "parallel",         # Support for parallel computation (often a base package but good to ensure)
    "devtools"          # Tools for package development (useful for installing from GitHub etc.)
)

$uniqueRPackages = $rPackages | Select-Object -Unique

$rExePath = $null
if ($ExplicitRScriptPath -and (Test-Path $ExplicitRScriptPath -PathType Leaf)) {
    if ($ExplicitRScriptPath -like "*Rscript.exe") {
        $rBinDir = Split-Path $ExplicitRScriptPath
        $potentialRExePath = Join-Path $rBinDir "R.exe"
        if (Test-Path $potentialRExePath -PathType Leaf) {
            $rExePath = $potentialRExePath
        }
    }
}

if (-not $rExePath) {
    $rCmdInfo = Get-Command R.exe -ErrorAction SilentlyContinue
    if ($rCmdInfo) {
        $rExePath = $rCmdInfo.Source
    }
}

if (-not $rExePath -or (-not (Test-Path $rExePath -PathType Leaf))) {
    Write-Error "R.exe not found via explicit path derivation or on system PATH. Cannot install R packages."
    Write-Warning "Ensure R is installed correctly and its 'bin' directory is added to the PATH environment variable."
    Write-Warning "You might need to restart your PowerShell session or computer."
    exit 1 
}

Write-Host "Using R.exe at: $rExePath" -ForegroundColor DarkGray
$cranMirror = "https://cran.rstudio.com/"

foreach ($pkg in $uniqueRPackages) {
    Write-Host "Checking/Installing R package: $pkg ..." -ForegroundColor Cyan
    
    $checkExpression = "if (!requireNamespace('$pkg', quietly = TRUE)) { quit(status = 10) } else { quit(status = 0) }"
    $checkCommand = """$rExePath"" CMD R -e ""$checkExpression""" 
    
    $exitCodeCheck = -1 
    try {
        Invoke-Expression $checkCommand
        $exitCodeCheck = $LASTEXITCODE
    } catch {
        Write-Warning "Error during package check for '$pkg': $($_.Exception.Message)"
    }

    if ($exitCodeCheck -eq 10) { 
        Write-Host "Package '$pkg' not found. Attempting to install..." -ForegroundColor Yellow
        
        $installExpression = "tryCatch({ install.packages('$pkg', dependencies=TRUE, repos='$cranMirror', Ncpus=max(1, parallel::detectCores()-1)) }, error = function(e) { print(paste('Error installing $pkg':, e)); quit(status = 1) })"
        $installCommand = """$rExePath"" CMD R -e ""$installExpression"""
        
        Write-Host "Executing: $installCommand" -ForegroundColor DarkGray
        $exitCodeInstall = -1
        try {
            Invoke-Expression $installCommand
            $exitCodeInstall = $LASTEXITCODE
        } catch {
            Write-Error "CRITICAL ERROR executing install command for '$pkg': $($_.Exception.Message)"
        }

        if ($exitCodeInstall -ne 0) {
            Write-Error "Failed to install R package: '$pkg'. Exit code from R: $exitCodeInstall."
        } else {
            Write-Host "Successfully installed R package: '$pkg'" -ForegroundColor Green
        }
    } elseif ($exitCodeCheck -eq 0) {
        Write-Host "R package '$pkg' is already installed." -ForegroundColor Green
    } else {
        Write-Warning "Could not definitively determine status of R package '$pkg'. Check exit code: $exitCodeCheck. Manual verification may be needed if installation fails."
    }
}

Write-Host "R Package installation process complete." -ForegroundColor Green