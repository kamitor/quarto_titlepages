
## 🚀 Getting Started: Installation & Setup

Choose the installation method appropriate for your operating system.

### Method 1: Automated Installer (Recommended for Windows Users)

This script will check for and help install Python, R, Quarto, required packages, LaTeX (TinyTeX), and the custom project font.

1.  **Download & Extract:**
    *   If you received this project as a ZIP file, extract its **entire contents** to a convenient folder on your computer (e.g., `C:\Users\YourName\Documents\ResilienceProject`).
    *   **Important:** Do not run any files directly from within the ZIP archive.

2.  **Run the Installer:**
    *   Navigate into the extracted project folder.
    *   Double-click the `run_installer.bat` file.
    *   A PowerShell window will open. Follow the on-screen prompts carefully.
    *   The script will guide you through downloading and installing any missing software. This process, especially the LaTeX (TinyTeX) installation, can take several minutes and requires an internet connection.

3.  **Post-Installation:**
    *   Once the script completes (ideally with success messages), your environment should be ready.
    *   A system restart or logging out/in might be necessary for the custom font to be fully recognized by all applications.

### Method 2: Manual Installation (For macOS, Linux, or Advanced Windows Users)

Ensure all components are correctly installed and added to your system's PATH.

1.  **Python (3.8+):**
    *   Install from [python.org](https://www.python.org/downloads/).
    *   Install required packages:
        ```bash
        pip install pandas pywin32
        ```
        *(Note: `pywin32` is Windows-specific, for Outlook integration).*

2.  **R (4.0+):**
    *   Install from [cran.r-project.org](https://cran.r-project.org/).
    *   Install required R packages (open R console):
        ```R
        install.packages(c("readr", "dplyr", "stringr", "tidyr", "ggplot2", "fmsb", "scales"))
        ```

3.  **Quarto CLI (1.3+):**
    *   Install from [quarto.org](https://quarto.org/docs/get-started/).

4.  **LaTeX Distribution:**
    *   **Recommended for Quarto:** Install TinyTeX by running `quarto install tinytex` in your terminal (may require admin rights).
    *   **Alternatively:** Install a full distribution like TeX Live (all platforms), MiKTeX (Windows), or MacTeX (macOS). Ensure XeLaTeX is included.

5.  **Custom Font (`QTDublinIrish.otf`):**
    *   Locate `QTDublinIrish.otf` in the `fonts/` directory of this project.
    *   Install it on your system:
        *   **Windows:** Right-click the file -> "Install" or "Install for all users".
        *   **macOS:** Double-click file -> "Install Font" in Font Book.
        *   **Linux:** Copy to `~/.fonts/` or `/usr/local/share/fonts/`, then run `fc-cache -fv`.
    *   A system restart or logout/login might be needed.

---

## ⚙️ How to Use

1.  **Navigate to Project Directory:**
    Open your terminal (Command Prompt, PowerShell, Bash, etc.) and change to the directory where you extracted/cloned the project files.
    ```bash
    cd path/to/your/ResilienceProject
    ```

2.  **Generate PDF Reports:**
    Run the report generation script:
    ```bash
    python generate_reports.py
    ```
    Generated PDF reports will be saved in the `reports/` subfolder.

3.  **Send Reports via Email (Windows with Outlook):**
    *   Ensure PDF reports have been generated first.
    *   **Configuration:** Open `send_emails.py` in a text editor:
        *   Set `TEST_MODE = False` for sending to actual recipient email addresses found in `cleaned_master.csv`.
        *   If `TEST_MODE = True`, emails will be sent to the address specified in `TEST_EMAIL`. Modify this as needed for your testing.
    *   Run the email script:
        ```bash
        python send_emails.py
        ```
    *   Microsoft Outlook must be running or able to be started by the script.

---

## 🛠️ Understanding Key Components

*   **`generate_reports.py`:** This Python script is the main engine for PDF creation. It reads `data/cleaned_master.csv`, iterates through each unique company, and calls the Quarto CLI to render the `example_3.qmd` template with company-specific parameters.
*   **`send_emails.py`:** This Python script (for Windows) uses `pywin32` to interact with Microsoft Outlook. It reads company and email data, finds corresponding PDFs in the `reports/` folder, and sends them as attachments.
*   **`example_3.qmd`:** This Quarto Markdown file is the blueprint for each PDF. It contains:
    *   YAML frontmatter for document metadata, LaTeX options, and custom page elements.
    *   R code chunks (``` `{r}` ```) to load data for the specific company, perform calculations (e.g., pillar scores), and generate plots (e.g., radar charts using `ggplot2` and `fmsb`).
    *   Markdown text which includes dynamic values from R using inline code (`r params$company`).
    *   LaTeX commands and environments for fine-grained control over PDF appearance.
*   **`data/cleaned_master.csv`:** The central CSV file containing all data points. Ensure this file is correctly formatted and populated for accurate report generation. The `company_name` column is crucial for identifying unique entities.

---

## ⚠️ Troubleshooting Common Issues

*   **Font Not Displaying Correctly in PDF:**
    *   Ensure `QTDublinIrish.otf` (from the `fonts/` folder) was installed correctly on your system.
    *   Try restarting your computer or logging out and back in, as LaTeX sometimes requires this to recognize new system fonts.
*   **LaTeX/PDF Generation Errors (e.g., "Error producing PDF"):**
    *   If using the Windows installer, re-run `run_installer.bat` and ensure TinyTeX installation (Step 6 in the script) completes successfully. This might require admin rights.
    *   Ensure your LaTeX distribution is complete and `xelatex` is available.
    *   Check the terminal output from `python generate_reports.py` for specific LaTeX error messages. These can often be searched online for solutions. If `example_3.qmd` has `keep-tex: true`, you can inspect the intermediate `.tex` file for clues.
*   **Python or R Package Errors (e.g., "ModuleNotFound", "package not found"):**
    *   If using the Windows installer, ensure it completed without errors for package installation.
    *   Manually install missing packages as described in the "Manual Installation" section.
    *   Check your internet connection and firewall/proxy settings if downloads fail.
*   **`send_emails.py` Fails (Windows/Outlook):**
    *   Confirm Microsoft Outlook is installed, configured with an email account, and preferably running.
    *   Ensure the `pywin32` Python package is installed.
*   **PowerShell Execution Policy (Windows Installer):**
    *   The `run_installer.bat` attempts to bypass the execution policy for its session. If it still fails, you might need to temporarily change your system's policy. Open PowerShell as Administrator and run `Set-ExecutionPolicy RemoteSigned -Scope Process -Force`, then try the `.bat` file again.

---

## 💡 Project Context & Future Development

This system was developed as part of the ResilienceScan initiative at Hogeschool Windesheim to streamline the creation and dissemination of strategic resilience profiles.

Potential future enhancements include:
*   More robust error handling and logging.
*   Configuration files for easier customization (e.g., email settings, output paths).
*   Support for additional output formats (e.g., HTML).
*   A web-based interface or dashboard for report management.

---

Thank you for using the Resilience Report Generator!