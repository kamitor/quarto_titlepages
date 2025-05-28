import os
import pandas as pd
import csv
from pathlib import Path
import shutil

# ✅ CONFIGURATION
ROOT = Path(__file__).resolve().parent
TEMPLATE = ROOT / "example_3.qmd"
DATA = ROOT / "data" / "cleaned_master.csv"
OUTPUT_DIR = ROOT / "reports"
COLUMN_MATCH = "company_name"

def load_csv(path):
    encodings = ["utf-8", "cp1252", "latin1"]
    for enc in encodings:
        try:
            with open(path, encoding=enc) as f:
                sample = f.read(2048)
                try:
                    sep = csv.Sniffer().sniff(sample).delimiter
                    print(f"✅ Delimiter '{sep}' with encoding '{enc}'")
                except Exception:
                    sep = ";"
                    print(f"⚠️ Using fallback delimiter ';' with encoding '{enc}'")
                return pd.read_csv(path, encoding=enc, sep=sep)
        except Exception as e:
            print(f"⚠️ Failed with encoding {enc}: {e}")
    raise RuntimeError("❌ Could not read CSV.")

def safe_filename(name):
    return "".join(c if c.isalnum() else "_" for c in str(name))

def generate_reports():
    df = load_csv(DATA)
    df.columns = df.columns.str.lower().str.strip()
    company_col = next((col for col in df.columns if COLUMN_MATCH in col), None)
    if not company_col:
        raise ValueError(f"❌ No column matching '{COLUMN_MATCH}'")

    companies = df[company_col].dropna().unique()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    for company in companies:
        safe_name = safe_filename(company)
        output_file = OUTPUT_DIR / f"{safe_name}.pdf"
        if output_file.exists():
            print(f"🔁 Skipping {company} (already exists)")
            continue

        print(f"📄 Generating: {company}")
        cmd = (
        f'quarto render "{TEMPLATE}" '
        f'-P company="{company}" '
        f'--to pdf '
        f'--output "{safe_name}.pdf"'
       )


        result = os.system(cmd)
        if result == 0:
            if Path(f"{safe_name}.pdf").exists():
                shutil.move(f"{safe_name}.pdf", output_file)
                print(f"✅ Saved: {output_file}")
            else:
                print(f"❌ Output file not found for {company}")
        else:
            print(f"❌ Failed: {company} (exit {result})")

if __name__ == "__main__":
    generate_reports()
