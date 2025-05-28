import os
import pandas as pd
import csv
from pathlib import Path

# ✅ CONFIG
ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "templates" / "resilience-report.qmd"
DATA = ROOT / "data" / "cleaned_master.csv"
OUTPUT_DIR = ROOT / "reports"
COMPLETED_PATH = ROOT / "completed.txt"
COLUMN_MATCH = "company_name"

def load_csv_resilient(path):
    if not path.exists():
        raise FileNotFoundError(f"❌ Data file not found at {path}")

    encodings = ["utf-8", "cp1252", "latin1"]
    for enc in encodings:
        try:
            with open(path, encoding=enc) as f:
                sample = f.read(2048)
                try:
                    sep = csv.Sniffer().sniff(sample).delimiter
                    print(f"✅ Detected delimiter '{sep}' with encoding '{enc}'")
                except Exception:
                    sep = ";"
                    print(f"⚠️ Could not detect delimiter — fallback to ';' with encoding '{enc}'")
                df = pd.read_csv(path, encoding=enc, sep=sep)
                print(f"✅ Loaded {len(df)} rows")
                return df
        except Exception as e:
            print(f"⚠️ Failed with encoding {enc}: {e}")
    raise ValueError("❌ Unable to load CSV with known encodings.")

def safe_filename(name):
    return "".join(c if c.isalnum() else "_" for c in str(name))

def generate_reports():
    df = load_csv_resilient(DATA)
    df.columns = df.columns.str.strip().str.lower()

    company_cols = [c for c in df.columns if COLUMN_MATCH in c]
    if not company_cols:
        raise ValueError(f"❌ Column containing '{COLUMN_MATCH}' not found.")

    company_col = company_cols[0]
    companies = df[company_col].dropna().unique()

    if os.path.exists(COMPLETED_PATH):
        with open(COMPLETED_PATH, "r", encoding="utf-8") as f:
            completed = set(line.strip() for line in f)
    else:
        completed = set()

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    for company in companies:
        safe_name = safe_filename(company)
        output_file = OUTPUT_DIR / f"{safe_name}.pdf"
        if safe_name in completed or output_file.exists():
            print(f"🔁 Skipping {company} (already exists)")
            continue

        print(f"📄 Generating report for: {company}")
        cmd = (
        f'quarto render "{TEMPLATE}" '
        f'-P company:"{company}" '
        f'--to pdf '
        f'--output-dir "{OUTPUT_DIR}" '
        f'--output "{safe_name}.pdf"'
        )

        result = os.system(cmd)
        if result == 0:
            with open(COMPLETED_PATH, "a", encoding="utf-8") as f:
                f.write(safe_name + "\n")
            f'--output "{safe_name}.pdf" --output-dir "{OUTPUT_DIR}"'
        else:
            print(f"❌ Failed for {company} (exit {result})")

if __name__ == "__main__":
    generate_reports()
