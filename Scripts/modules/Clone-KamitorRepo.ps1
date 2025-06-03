# Script: modules/Clone-KamitorRepo.ps1
# Purpose: Clones the 'kamitor/quarto_titlepages' GitHub repository.

Write-Host "Starting cloning of 'kamitor/quarto_titlepages' repository..." -ForegroundColor Yellow

# These variables are expected to be set by Main-Installer.ps1
# $Global:RepoUrl = "https://github.com/kamitor/quarto_titlepages.git"
# $Global:FullClonePath = "$HOME\Documents\GitHub\kamitor_quarto_titlepages"
# $Global:CloneParentDir = "$HOME\Documents\GitHub"

if (-not ($Global:RepoUrl) -or -not ($Global:FullClonePath) -or -not ($Global:CloneParentDir)) {
    Write-Error "Required global variables (RepoUrl, FullClonePath, CloneParentDir) are not set."
    exit 1
}

Write-Host "Repository URL: $($Global:RepoUrl)"
Write-Host "Target clone path: $($Global:FullClonePath)"
Write-Host "Parent directory for clone: $($Global:CloneParentDir)"

# Check if Git is available
if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    Write-Error "Git command not found. Cannot clone repository."
    exit 1
}
Write-Host "Found Git: $((git --version).Trim())"

try {
    # Create the parent directory if it doesn't exist
    if (-not (Test-Path $Global:CloneParentDir -PathType Container)) {
        Write-Host "Parent directory '$($Global:CloneParentDir)' does not exist. Creating it..."
        try {
            New-Item -ItemType Directory -Path $Global:CloneParentDir -Force -ErrorAction Stop | Out-Null
            Write-Host "Parent directory '$($Global:CloneParentDir)' created successfully." -ForegroundColor Green
        } catch {
            Write-Error "Failed to create parent directory '$($Global:CloneParentDir)'. Error: $($_.Exception.Message)"
            exit 1
        }
    } else {
        Write-Host "Parent directory '$($Global:CloneParentDir)' already exists."
    }

    # Check if the repository already exists (e.g., a .git folder inside)
    if (Test-Path (Join-Path -Path $Global:FullClonePath -ChildPath ".git")) {
        Write-Warning "Repository already seems to be cloned at '$($Global:FullClonePath)'. Skipping clone."
        # Consider adding logic here to pull latest if it exists, or remove and re-clone,
        # but for now, just skipping to prevent errors.
        exit 0
    }
    # Also check if it's a non-empty directory that's NOT a git repo
    if (Test-Path $Global:FullClonePath -PathType Container) {
        if (Get-ChildItem -Path $Global:FullClonePath) { # Check if directory has any items
             Write-Warning "Directory '$($Global:FullClonePath)' exists and is not empty, but not a git repo. Skipping clone to avoid data loss."
             exit 1 # Indicate an issue that prevents cloning
        }
    }


    Write-Host "Executing: git clone $($Global:RepoUrl) '$($Global:FullClonePath)'"
    git clone $Global:RepoUrl $Global:FullClonePath
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Git clone command failed with exit code $LASTEXITCODE."
        exit 1
    }

    Write-Host "Repository cloned successfully to '$($Global:FullClonePath)'." -ForegroundColor Green
    exit 0
} catch {
    Write-Error "An unexpected error occurred during repository cloning: $($_.Exception.Message)"
    exit 1
}