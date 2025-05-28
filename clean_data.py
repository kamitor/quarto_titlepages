import pandas as pd
import csv

source_path = "./data/Resilience - MasterDatabase(MasterData).csv"
output_path = "./data/cleaned_master.csv"

# Try to load the file with encoding and fallback delimiter
def clean_and_save():
    encodings = ["utf-8", "cp1252", "latin1"]
    for enc in encodings:
        try:
            with open(source_path, encoding=enc) as f:
                lines = list(csv.reader(f, delimiter=";"))

                print(f"✅ File loaded with encoding {enc}")

                # Print first 5 lines to debug structure
                for i in range(5):
                    print(f"Row {i+1}:", lines[i])

                # Assume row 2 (index 1) is header, rest is data
                header_row = lines[1]
                data_rows = lines[2:]

                df = pd.DataFrame(data_rows, columns=header_row)

                # Clean up column names
                df.columns = (
                    pd.Series(df.columns)
                    .str.strip()
                    .str.lower()
                    .str.replace(" ", "_")
                    .str.replace(r"[^\w]", "", regex=True)
                )

                # Save to cleaned CSV
                df.to_csv(output_path, index=False, encoding="utf-8")
                print(f"✅ Cleaned file saved to: {output_path}")
                return
        except Exception as e:
            print(f"❌ Failed with encoding {enc}: {e}")
    print("❌ Could not process the file.")

if __name__ == "__main__":
    clean_and_save()
