import pandas as pd
import os

# File paths
EXCEL_FILE = os.path.expanduser("~/Documents/CMS_Decided_16-12-2025 17_26_11 (1).xlsx")
FILLED_ENTRIES = os.path.expanduser("~/Documents/filled_entries.txt")

def extract_excel_to_filled():
    """Extract rows 1-787 from Excel and write to filled_entries.txt WITH NUMBERING"""
    
    print("Reading Excel file...")
    df = pd.read_excel(EXCEL_FILE, sheet_name='Sheet1')
    
    # Extract rows 1 to 787 (Python uses 0-based indexing, so 0:787)
    rows = df.iloc[611:787]
    
    print(f"✓ Found {len(rows)} rows to write")
    
    # Open filled_entries.txt and write with numbering
    with open(FILLED_ENTRIES, 'w', encoding='utf-8') as f:
        for idx, (_, row) in enumerate(rows.iterrows(), start=1):
            # Convert row to tab-separated string
            row_data = '\t'.join(str(val) for val in row.values)
            # Write with number at start
            f.write(f"{idx}\t{row_data}\n")
    
    print(f"✓ Wrote {len(rows)} entries with numbering to {FILLED_ENTRIES}")

if __name__ == "__main__":
    extract_excel_to_filled()