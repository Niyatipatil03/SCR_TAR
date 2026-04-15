# =========================================================
# SCR & TAR MERGE SCRIPT
# Incremental by FILE DATE
# =========================================================

import os
import re
import pandas as pd
from datetime import datetime, date

# ================= PATHS =================
BASE_PATH    = r"C:\Users\25033810\OneDrive - Mahindra & Mahindra Ltd\Field Quality Nasik - SCR & TAR"
SCR_FOLDER   = os.path.join(BASE_PATH, "Daily SCR")
TAR_FOLDER   = os.path.join(BASE_PATH, "Daily TAR")
COMPILED_FILE = os.path.join(BASE_PATH, "SCR_TAR_Compiled.xlsx")

print("📂 BASE PATH  :", BASE_PATH)
print("📄 COMPILED   :", COMPILED_FILE)

# ================= COLUMN STANDARDIZATION =================
COLUMN_MAPPING = {
    "TAR No."                 : "Report No.",
    "Report No."              : "Report No.",
    "Serial No"               : "Vehicle Sr No",
    "Vehicle Sr No"           : "Vehicle Sr No",
    "Model Family"            : "Model Family",
    "Created Date"            : "Created date",
    "Date"                    : "Created date",
    "Short Description"       : "Description",
    "Description of Complaint": "Description",
    "Description"             : "Description",
}

SCR_COLUMNS_TO_DROP = [
    "Status", "Cotek", "SCR Type", "Engine No.",
    "Additional Vehicles", "Date of Failure",
]

TAR_COLUMNS_TO_DROP = [
    "Verbatim Code", "Registration No.", "Created By",
    "Conversations", "Closed Date",
]

# ================= VSN MAPS =================
VSN_YEAR_MAP = {
    "N": 2022, "P": 2023, "R": 2024,
    "S": 2025, "T": 2026, "U": 2027,
    "V": 2028, "W": 2029,
}

VSN_MONTH_MAP = {
    "A": 1,  "B": 2,  "C": 3,  "D": 4,
    "E": 5,  "F": 6,  "G": 7,  "H": 8,
    "J": 9,  "K": 10, "L": 11, "M": 12,
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
    """
    Try DD-MMM-YY first (e.g. 14-Apr-26), then DDMMYY (e.g. 140426).
    Returns a date object or None.
    """
    fname_up = fname.upper()

    # Pattern 1: DD-MMM-YY  (most specific — try first)
    m = re.search(r"(\d{2})-([A-Z]{3})-(\d{2})", fname_up)
    if m:
        try:
            return datetime.strptime(m.group(0), "%d-%b-%y").date()
        except ValueError:
            pass

    # Pattern 2: exactly 6 consecutive digits DDMMYY
    m = re.search(r"(?<!\d)(\d{6})(?!\d)", fname_up)
    if m:
        try:
            return datetime.strptime(m.group(1), "%d%m%y").date()
        except ValueError:
            pass

    return None


def compute_vsn_date(vsn):
    """
    Extract manufacture year + month from Vehicle Sr No.
    Format assumption: char[0] = year code, char[2] = month code.
    e.g. T2B... → T=2026, B=Feb → 2026-02-01
    Returns pd.Timestamp or NaT.
    """
    try:
        vsn = str(vsn).strip().upper()
        if len(vsn) < 3:
            return pd.NaT
        year  = VSN_YEAR_MAP.get(vsn[0])
        month = VSN_MONTH_MAP.get(vsn[2])
        if year and month:
            return pd.Timestamp(year=year, month=month, day=1)
    except Exception:
        pass
    return pd.NaT


# ================= LOAD EXISTING COMPILED FILE =================
if os.path.exists(COMPILED_FILE):
    compiled_df = pd.read_excel(COMPILED_FILE, engine="openpyxl")
    if "File_Date" in compiled_df.columns:
        max_file_date = (
            pd.to_datetime(compiled_df["File_Date"], errors="coerce")
            .dt.date.max()
        )
    else:
        max_file_date = None
else:
    compiled_df   = pd.DataFrame()
    max_file_date = None

print("📅 Max FILE date already merged:", max_file_date)

incoming_data = []


# ================= PROCESS ONE FOLDER =================
def process_folder(folder, folder_name):
    print(f"\n📂 Scanning {folder_name} folder: {folder}")
    today = date.today()

    for f in sorted(os.listdir(folder)):
        if not f.lower().endswith(".xlsx") or f.startswith("~$"):
            continue

        file_date = extract_date_from_filename(f)
        print(f"   → {f}  |  Date: {file_date}")

        if file_date is None:
            print("      ⚠️  Skipped (no date found in filename)")
            continue

        if max_file_date and file_date <= max_file_date:
            print("      ⏭️  Skipped (already merged)")
            continue

        # ---- Read & standardize ----
        df = pd.read_excel(os.path.join(folder, f), engine="openpyxl")
        df = standardize_columns(df)

        # ---- Drop unwanted columns ----
        drop_list = SCR_COLUMNS_TO_DROP if folder_name == "SCR" else TAR_COLUMNS_TO_DROP
        df.drop(columns=[c for c in drop_list if c in df.columns],
                inplace=True, errors="ignore")

        # ---- Ensure VSN Date column ----
        if "VSN Date" not in df.columns and "Vehicle Sr No" in df.columns:
            df["VSN Date"] = df["Vehicle Sr No"].apply(compute_vsn_date)
        df["VSN Date"] = pd.to_datetime(df.get("VSN Date"), errors="coerce")

        # ---- Ensure Days column (days from VSN Date to today) ----
        if "Days" not in df.columns:
            df["Days"] = df["VSN Date"].apply(
                lambda d: (today - d.date()).days if pd.notna(d) else None
            )

        # ---- Tag rows ----
        df["Source_Folder"] = folder_name
        df["File_Date"]     = file_date

        incoming_data.append(df)
        print(f"      ✅ MERGED {folder_name}: {f}  ({len(df)} rows)")


# ================= RUN =================
process_folder(SCR_FOLDER, "SCR")
process_folder(TAR_FOLDER, "TAR")

if not incoming_data:
    print("\n⚠️  No new files detected — nothing merged.")
else:
    incoming_df = pd.concat(incoming_data, ignore_index=True)
    final_df    = pd.concat([compiled_df, incoming_df], ignore_index=True)
    final_df.to_excel(COMPILED_FILE, index=False)

    print("\n✅ MERGE COMPLETE")
    print(f"   Rows added : {len(incoming_df)}")
    print(f"   Total rows : {len(final_df)}")
