import os
import re
import pandas as pd
import pdfplumber
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# --- CONFIGURATION / SETTINGS ---
# You can easily adapt the script to different projects by changing these keys.
CONFIG = {
    "SOURCE_EXCEL": 'data_source.xlsx',
    "DOCUMENTS_FOLDER": 'pdf_vault',
    "OUTPUT_FILE": 'synchronized_result.xlsx',
    "HIGHLIGHT_COLOR": "FFFF00", # Yellow
    "MATCH_KEYS": {
        "key1": "ADA",      # Change this to your Excel's first key column
        "key2": "PARSEL",   # Change this to your Excel's second key column
        "target": "YIL"     # The column you want to fill/update
    },
    "REGEX_PATTERNS": {
        "k1": r'Block.*?(\d+)',           # Pattern to find first key in PDF
        "k2": r'Parcel.*?(\d+)',          # Pattern to find second key in PDF
        "value": r'Year.*?(\d{4})'        # Pattern to find the actual value
    }
}

def extract_data_from_pdf(file_path):
    """Parses PDF text and extracts key-value pairs using Regex."""
    extracted_items = []
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text: continue
                text = text.replace('\n', ' ')
                
                # Regex matching
                p = CONFIG["REGEX_PATTERNS"]
                m1 = re.search(p["k1"], text, re.IGNORECASE)
                m2 = re.search(p["k2"], text, re.IGNORECASE)
                mv = re.search(p["value"], text, re.IGNORECASE)
                
                if m1 and m2 and mv:
                    extracted_items.append({
                        'k1': int(m1.group(1)),
                        'k2': int(m2.group(1)),
                        'val': int(mv.group(1))
                    })
    except Exception as e:
        print(f"Read Error: {e}")
    return extracted_items

def run_sync_process():
    # 1. Load and Standardize Excel
    df = pd.read_excel(CONFIG["SOURCE_EXCEL"])
    df.columns = [str(c).strip().upper() for c in df.columns]
    
    keys = CONFIG["MATCH_KEYS"]
    
    # 2. Build PDF Database
    pdf_db = {}
    if os.path.exists(CONFIG["DOCUMENTS_FOLDER"]):
        files = [f for f in os.listdir(CONFIG["DOCUMENTS_FOLDER"]) if f.endswith('.pdf')]
        for file in files:
            path = os.path.join(CONFIG["DOCUMENTS_FOLDER"], file)
            records = extract_data_from_pdf(path)
            for r in records:
                pdf_db[(r['k1'], r['k2'])] = r['val']

    # 3. Synchronize Data
    updated_indices = []
    for idx, row in df.iterrows():
        # Match Excel row with PDF database using keys
        match_result = pdf_db.get((row[keys['key1']], row[keys['key2']]))
        
        if match_result and pd.isna(row[keys['target']]):
            df.at[idx, keys['target']] = match_result
            updated_indices.append(idx)

    # 4. Save and Format
    df.to_excel(CONFIG["OUTPUT_FILE"], index=False)
    print(f"Sync complete. {len(updated_indices)} rows updated.")

if __name__ == "__main__":
    run_sync_process()