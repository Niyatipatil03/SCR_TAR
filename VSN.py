# =========================================================
# SCR & TAR MERGE SCRIPT
# Incremental by FILE DATE (CORRECTED)
# =========================================================

import os
import re
import pandas as pd
from datetime import datetime

# ================= ✅ CORRECT PATH =================
BASE_PATH = r"C:\Users\25033810\OneDrive - Mahindra & Mahindra Ltd\Field Quality Nasik - SCR & TAR"
SCR_FOLDER = os.path.join(BASE_PATH, "Daily SCR")
TAR_FOLDER = os.path.join(BASE_PATH, "Daily TAR")
COMPILED_FILE = os.path.join(BASE_PATH, "SCR_TAR_Compiled.xlsx")

print("📂 BASE PATH:", BASE_PATH)
print("📄 COMPILED FILE:", COMPILED_FILE)

# ================= COLUMN STANDARDIZATION =================
COLUMN_MAPPING = {
    "Report No.": "Report No.",
    "TAR No.": "Report No.",
    "Serial No": "Vehicle Sr No",
    "Vehicle Sr No": "Vehicle Sr No",
    "Job card No": "Model Family",
    "Model Family": "Model Family",
    "Created Date": "Created date",
    "Date": "Created date",
    "Short Description": "Description",
    "Description of Complaint": "Description",
    "Description": "Description",
}

SCR_COLUMNS_TO_DROP = [
    "Status", "Cotek", "SCR Type", "Engine No.",
    "Additional Vehicles", "Date of Failure"
]

TAR_COLUMNS_TO_DROP = [
    "Verbatim Code", "Registration No.", "Created By",
    "Conversations", "Closed Date"
]

VSN_YEAR_MAP = {
    "N": 2022, "P": 2023, "R": 2024, "S": 2025,
    "T": 2026, "U": 2027, "V": 2028, "W": 2029
}

VSN_MONTH_MAP = {
    "A": 1, "B": 2, "C": 3, "D": 4,
    "E": 5, "F": 6, "G": 7, "H": 8,
    "J": 9, "K": 10, "L": 11, "M": 12
}

# ================= HELPERS =================
def standardize_columns(df):
    df.columns = (
        df.columns.astype(str)
        .str.replace("\u00A0", " ", regex=False)
        .str.strip()
    )
    df.rename(columns=COLUMN_MAPPING, inplace=True)
    return df

def extract_date_from_filename(fname):
    fname = fname.upper()

    m1 = re.search(r"(\d{2})(\d{2})(\d{2})", fname)
    if m1:
        return datetime.strptime(m1.group(0), "%d%m%y").date()

    m2 = re.search(r"(\d{2})-([A-Z]{3})-(\d{2})", fname)
    if m2:
        return datetime.strptime(m2.group(0), "%d-%b-%y").date()

    return None

# ================= LOAD EXISTING (FILE_DATE) =================
if os.path.exists(COMPILED_FILE):
    compiled_df = pd.read_excel(COMPILED_FILE, engine="openpyxl")
    if "File_Date" in compiled_df.columns:
        max_file_date = pd.to_datetime(compiled_df["File_Date"], errors="coerce").dt.date.max()
    else:
        max_file_date = None
else:
    compiled_df = pd.DataFrame()
    max_file_date = None

print("📅 Max FILE date already merged:", max_file_date)

incoming_data = []

# ================= PROCESS FOLDERS =================
def process_folder(folder, folder_name):
    print(f"\n📂 Scanning {folder_name} folder:", folder)

    for f in os.listdir(folder):
        if not f.lower().endswith(".xlsx") or f.startswith("~$"):
            continue

        file_date = extract_date_from_filename(f)
        print("   → Found file:", f, "| Date:", file_date)

        if not file_date:
            continue

        if max_file_date and file_date <= max_file_date:
            print("     ⏭️ Skipped (already merged)")
            continue

        df = pd.read_excel(os.path.join(folder, f), engine="openpyxl")
        df = standardize_columns(df)

        if folder_name == "SCR":
            df.drop(columns=[c for c in SCR_COLUMNS_TO_DROP if c in df.columns], inplace=True, errors="ignore")
        else:
            df.drop(columns=[c for c in TAR_COLUMNS_TO_DROP if c in df.columns], inplace=True, errors="ignore")

        df["Source_Folder"] = folder_name
        df["File_Date"] = file_date
        incoming_data.append(df)

        print(f"     ✅ MERGED {folder_name}: {f} ({len(df)} rows)")

process_folder(SCR_FOLDER, "SCR")
process_folder(TAR_FOLDER, "TAR")

if not incoming_data:
    print("⚠️ No new files detected. Nothing merged.")
    exit()

incoming_df = pd.concat(incoming_data, ignore_index=True)

# ================= FINAL SAVE =================
final_df = pd.concat([compiled_df, incoming_df], ignore_index=True)
final_df.to_excel(COMPILED_FILE, index=False)

print("\n✅ MERGE COMPLETE")
print("✅ Rows added:", len(incoming_df))
print("✅ Total rows:", len(final_df))
print("✅ 13‑Apr & 14‑Apr WILL NOW BE PRESENT")