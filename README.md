# 📊 Resilience Scan Report Generator

This project automates the generation of PDF reports for multiple companies based on their data using a single Quarto template.

---

## ✅ Project Structure

```
.
├── example_3.qmd              # Quarto template for the report
├── generate_all_reports.py    # Python script to render all company reports
├── data/
│   └── cleaned_master.csv     # Input data file with company details
├── reports/                   # Output folder for generated PDFs
├── img/                       # Images used in the PDF cover/title
│   ├── logo.png
│   ├── corner-bg.png
│   └── otter-bar.jpeg
├── tex/
│   ├── dedication.tex         # Custom dedication page
│   └── copyright.tex          # Copyright/license notice
```

---

## ⚙️ How It Works

1. The Python script:
   - Loads `cleaned_master.csv`
   - Iterates through all unique company names
   - Renders a separate PDF report per company using the Quarto template
   - Saves each report in the `reports/` folder
   - Skips reports already generated

2. The Quarto template:
   - Accepts a parameter `company`
   - Filters data based on that company name
   - Renders a personalized PDF for each company
   - Includes a custom cover page, title page, and formatted design

---

## 🧪 Features Tested

- Parameterized Quarto rendering using `-P company:"<Name>"`
- CSV reading with flexible encodings
- Clean file naming for output
- PDF generation via XeLaTeX
- Echo, warnings, and messages disabled in final output
- Custom cover page and title page design using `quarto-titlepages`

---

## 💻 Install Guide

### 1. Install Python 3

https://www.python.org/downloads/  
Make sure `python` and `pip` are in your PATH.

### 2. Install Required Python Packages

```bash
pip install pandas
```

### 3. Install Quarto

https://quarto.org/docs/get-started/

Make sure `quarto` is accessible from the command line.

### 4. Install LaTeX

Install a full TeX distribution that includes XeLaTeX:

- **Windows**: [TeX Live](https://tug.org/texlive/windows.html) or [MiKTeX](https://miktex.org/)
- **Linux**: `sudo apt install texlive-full`
- **Mac**: [MacTeX](https://tug.org/mactex/)

---

## 🏁 Run

To generate the reports:

```bash
python generate_all_reports.py
```

---

## 🚧 Next Steps

## ✅ Progress Checklist: Quarto Report Automation System

### 🟩 Core Functionality (Done)
- [x] Created Data-Clean Function for CSV
- [x] Load cleaned company data from `data/cleaned_master.csv`
- [x] Generate individual PDF reports per company using Quarto and parameters
- [x] Use custom LaTeX-styled `example_3.qmd` template with branding
- [x] Store generated reports in the `reports/` folder
- [x] Skip already generated reports to save time

---

### 🔧 Technical Improvements (Planned)
- [ ] Add basic error handling for missing/malformed CSV data
- [ ] Refactor script for better retry support and clearer logging
- [ ] Parametrize output formats (PDF, HTML)
- [ ] Automatically move output file instead of relying on Quarto to put it in the right location

---

### 📊 Report Content Enhancements (Planned)
- [ ] Add summary statistics per company (e.g., score, compliance level, key metrics)
- [ ] Add basic plots (e.g., bar chart of key indicators)
- [ ] Ensure template gracefully handles missing data per company
- [ ] Modularize template to adapt to future layout/styling changes

---

### 📤 Automation Pipeline (In Progress)
- [ ] Integrate Outlook for automatic email delivery of each report
- [ ] Add `.env` or config file for Outlook credentials and recipients
- [ ] Log all sent emails and failures to a local or shared log file
- [ ] Allow optional delay between emails to avoid rate-limiting

---

### 🧪 Future Readiness (Stretch Goals)
- [ ] Allow support for multiple rows per company (grouped aggregation)
- [ ] Add tagging or categorization logic (e.g., “High Risk”, “Compliant”, etc.)
- [ ] Publish summary dashboard (Quarto HTML or Streamlit) showing company status
- [ ] Use GitHub Actions or cron job to re-generate reports weekly/monthly

---

### 🚨 Watch-Outs
- [ ] Ensure the `example_3.qmd` template always includes a fallback if no data is found
- [ ] Avoid reusing output names that could clash (sanitize filenames carefully)
- [ ] Monitor `.quarto` folder lock issues (XeLaTeX temp files)

---

Built with ❤️ for the ResilienceScan project at Hogeschool Windesheim.
