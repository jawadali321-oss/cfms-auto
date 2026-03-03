import pandas as pd
from pathlib import Path

files = [
    "CMS_Decided_16-12-2025 17_26_11 (1).xlsx",
    "Decision Sep.2025 Rehan ul hasan sb..xlsx",
    "khadim shb uplod.xlsx",
    "Rehan Ul Hassan _ oct 2025.xlsx",
    "Rehan Ul Hassan _ Tariq Mahmood Kahut Disposal October 2025.xlsx",
    "zameer hussain nov.xlsx",
    "Zamir Hussain Disposal October (6).xlsx"
]

keywords = [
    "Agreed",
    "منظور شد",
    "فیصلہ شد",
    "فیصلہ شدمنظور شد"
]

output = Path("cancellation.txt")
output.write_text("", encoding="utf-8")

for file in files:
    try:
        xls = pd.ExcelFile(file)
        last_sheet = xls.sheet_names[-1]
        df = pd.read_excel(file, sheet_name=last_sheet, dtype=str)

        for _, row in df.iterrows():
            row_text = " ".join(row.fillna("").astype(str))
            if any(k in row_text for k in keywords):
                with output.open("a", encoding="utf-8") as f:
                    f.write(f"{file} | {last_sheet}\n")
                    f.write(row_text + "\n")
                    f.write("-" * 80 + "\n")

    except Exception as e:
        with output.open("a", encoding="utf-8") as f:
            f.write(f"ERROR in {file}: {e}\n")
