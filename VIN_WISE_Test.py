# =========================================================
# DAILY SCR & TAR MAIL SCRIPT
# - Sends ONE mail per run (for yesterday's data)
# - Filters:
#       Created date  = yesterday
#       VSN Year      = current calendar year  (dynamic)
#       Days          <= 210
# =========================================================

import pandas as pd
import win32com.client as win32
from datetime import date, timedelta

# ================= CONFIG =================
COMPILED_FILE = (
    r"C:\Users\25033810\OneDrive - Mahindra & Mahindra Ltd"
    r"\Field Quality Nasik - SCR & TAR\SCR_TAR_Compiled.xlsx"
)

MAIL_TO = "Patil.Niyati@mahindra.com"
MAIL_CC = ""          # add CC addresses separated by ";"

# ================= VSN MAPS (for dynamic year detection) =================
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

# ================= COLUMN NAMES =================
DATE_COL     = "Created date"
DATE_PARSED  = "Created date(parsed)"
VSN_DATE_COL = "VSN Date"
DAYS_COL     = "Days"
REPORT_COL   = "Report No."
MODEL_COL    = "Model Family"
VEHICLE_COL  = "Vehicle Sr No"
DESC_COL     = "Description"
DEALER_COL   = "Dealer Name"
CITY_COL     = "Dealer City"
KRSS_COL     = "Krss covered"

# Ordered model buckets for summary table
MODEL_BUCKETS = ["XUV", "NEW Thar", "Scorpio Classic", "Thar ROXX"]


# ================= HELPERS =================

def compute_vsn_date(vsn):
    """Derive manufacture year+month from Vehicle Sr No.
    char[0] = year code  (R=2024, S=2025, T=2026 …)
    char[2] = month code (A=Jan, B=Feb, E=May …)
    e.g. S2E90733 → S=2025, E=May → 2025-05-01
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


def detect_type(report_no):
    r = str(report_no).upper().strip()
    if r.startswith("S-"):
        return "SCR"
    if r.startswith("T-"):
        return "TAR"
    return "UNKNOWN"


def get_model_bucket(model):
    """Map detailed model names to summary-table bucket."""
    m = str(model).upper()
    if "THAR ROXX" in m or "THAR ROX" in m:
        return "Thar ROXX"
    if "NEW THAR" in m:
        return "NEW Thar"
    if "SCORPIO CLASSIC" in m:
        return "Scorpio Classic"
    if "XUV" in m:
        return "XUV"
    return str(model)   # keep original if no bucket matched


# -------- HTML table builders --------

_HDR  = ('style="background-color:#1F3864;color:#FFFFFF;font-weight:bold;'
         'padding:6px 10px;text-align:left;white-space:nowrap;'
         'font-family:Calibri;font-size:10pt;"')
_HDR_C = ('style="background-color:#1F3864;color:#FFFFFF;font-weight:bold;'
          'padding:6px 12px;text-align:center;white-space:nowrap;'
          'font-family:Calibri;font-size:10pt;"')
_CELL  = 'style="padding:5px 10px;border-bottom:1px solid #E0E0E0;font-family:Calibri;font-size:10pt;"'
_CELL_C = 'style="padding:5px 12px;border-bottom:1px solid #E0E0E0;text-align:center;font-family:Calibri;font-size:10pt;"'


def _data_table(df, columns):
    """Render a styled HTML table; columns not present in df are skipped."""
    cols = [c for c in columns if c in df.columns]
    if not cols:
        return "<p style='font-family:Calibri'>No data to display.</p>"

    header_html = "".join(f"<th {_HDR}>{c}</th>" for c in cols)
    rows_html = ""
    for i, (_, row) in enumerate(df[cols].iterrows()):
        bg = "#F5F7FA" if i % 2 == 0 else "#FFFFFF"
        cells = "".join(
            f'<td {_CELL}>{("" if pd.isna(v) else v)}</td>'
            for v in row.values
        )
        rows_html += (
            f'<tr style="background-color:{bg};">{cells}</tr>\n'
        )

    return (
        f'<table style="border-collapse:collapse;width:100%;'
        f'font-family:Calibri;font-size:10pt;">'
        f'<thead><tr>{header_html}</tr></thead>'
        f'<tbody>{rows_html}</tbody>'
        f'</table>'
    )


def _summary_table(scr_df, tar_df):
    """Model-family bucketed summary: rows = buckets, cols = SCR | TAR."""
    def bucket_counts(df):
        if df.empty:
            return {}
        return df[MODEL_COL].apply(get_model_bucket).value_counts().to_dict()

    scr_counts = bucket_counts(scr_df)
    tar_counts = bucket_counts(tar_df)

    # Use MODEL_BUCKETS order first, then append any unlisted buckets
    all_keys = list(dict.fromkeys(
        MODEL_BUCKETS
        + list(scr_counts.keys())
        + list(tar_counts.keys())
    ))
    active = [b for b in all_keys if scr_counts.get(b, 0) + tar_counts.get(b, 0) > 0]

    rows_html = ""
    total_scr = total_tar = 0
    for i, bucket in enumerate(active):
        s = scr_counts.get(bucket, 0)
        t = tar_counts.get(bucket, 0)
        total_scr += s
        total_tar += t
        bg = "#F5F7FA" if i % 2 == 0 else "#FFFFFF"
        rows_html += (
            f'<tr style="background-color:{bg};">'
            f'<td {_CELL}>{bucket}</td>'
            f'<td {_CELL_C}>{s}</td>'
            f'<td {_CELL_C}>{t}</td>'
            f'</tr>\n'
        )
    # TOTAL row
    rows_html += (
        f'<tr style="background-color:#D9E1F2;font-weight:bold;">'
        f'<td {_CELL}>TOTAL</td>'
        f'<td {_CELL_C}>{total_scr}</td>'
        f'<td {_CELL_C}>{total_tar}</td>'
        f'</tr>'
    )

    return (
        f'<table style="border-collapse:collapse;font-family:Calibri;font-size:10pt;">'
        f'<thead><tr>'
        f'<th {_HDR}>Model Family</th>'
        f'<th {_HDR_C}>SCR</th>'
        f'<th {_HDR_C}>TAR</th>'
        f'</tr></thead>'
        f'<tbody>{rows_html}</tbody>'
        f'</table>'
    )


def send_mail(subject, html_body):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To      = MAIL_TO
    mail.CC      = MAIL_CC
    mail.Subject = subject
    mail.HTMLBody = html_body
    mail.Send()


# ================= MAIN =================

def main():

    # ── Load compiled file ──
    try:
        df = pd.read_excel(COMPILED_FILE, engine="openpyxl")
    except PermissionError:
        raise RuntimeError(
            "SCR_TAR_Compiled.xlsx is open or locked. "
            "Close Excel / wait for OneDrive sync, then rerun."
        )

    # ── Normalize column names ──
    df.columns = (
        df.columns.astype(str)
        .str.replace("\u00A0", " ", regex=False)
        .str.strip()
    )

    # ── Parse Created date ──
    if DATE_PARSED not in df.columns:
        df[DATE_PARSED] = pd.to_datetime(
            df[DATE_COL], errors="coerce", dayfirst=True
        ).dt.date
    df[DATE_PARSED] = pd.to_datetime(df[DATE_PARSED], errors="coerce").dt.date

    # ── Ensure VSN Date (compute from Vehicle Sr No if absent) ──
    if VSN_DATE_COL not in df.columns:
        df[VSN_DATE_COL] = df[VEHICLE_COL].apply(compute_vsn_date)
    df[VSN_DATE_COL] = pd.to_datetime(df[VSN_DATE_COL], errors="coerce")

    # ── Ensure Days column = Created date − VSN Date (vehicle age when issue raised) ──
    if DAYS_COL not in df.columns:
        created_ts = pd.to_datetime(df[DATE_COL], errors="coerce", dayfirst=True)
        df[DAYS_COL] = (created_ts - df[VSN_DATE_COL]).dt.days
    df[DAYS_COL] = pd.to_numeric(df[DAYS_COL], errors="coerce")

    # ── Type detection ──
    df["Type"] = df[REPORT_COL].apply(detect_type)

    # ── Dynamic VSN year = current calendar year ──
    current_year = date.today().year        # e.g. 2026

    # ── Target date = YESTERDAY only ──
    target_date = date.today() - timedelta(days=1)

    # ── Apply filters ──
    filtered = df[
        (df[DATE_PARSED]    == target_date)    &
        (df[DAYS_COL]       <= 210)            &
        (df[VSN_DATE_COL].dt.year == current_year)
    ].copy()

    if filtered.empty:
        print(
            f"ℹ️  No data for {target_date} "
            f"(VSN year {current_year}, Days ≤ 210). Mail NOT sent."
        )
        return

    scr_df = filtered[filtered["Type"] == "SCR"].reset_index(drop=True)
    tar_df = filtered[filtered["Type"] == "TAR"].reset_index(drop=True)

    # ── Date strings for mail ──
    target_str  = target_date.strftime("%d-%b-%Y")

    def latest_str(sub):
        if sub.empty:
            return target_str
        latest = sub[DATE_PARSED].max()
        return latest.strftime("%d-%b-%Y") if latest else target_str

    scr_latest_str = latest_str(scr_df)
    tar_latest_str = latest_str(tar_df)

    # ── Buckets present (for summary header) ──
    def active_buckets(sub):
        if sub.empty:
            return []
        return list(dict.fromkeys(
            sub[MODEL_COL].apply(get_model_bucket).tolist()
        ))

    all_active  = list(dict.fromkeys(active_buckets(scr_df) + active_buckets(tar_df)))
    # Keep MODEL_BUCKETS order; append any extra
    ordered     = [b for b in MODEL_BUCKETS if b in all_active]
    ordered    += [b for b in all_active if b not in ordered]
    buckets_str = ", ".join(ordered) if ordered else "N/A"

    # ── Build HTML tables ──
    summary_html = _summary_table(scr_df, tar_df)

    SCR_COLS = [REPORT_COL, DATE_COL, DEALER_COL, CITY_COL,
                MODEL_COL, VEHICLE_COL, DESC_COL, KRSS_COL]
    TAR_COLS = ["Type", REPORT_COL, DATE_COL,
                MODEL_COL, VEHICLE_COL, DESC_COL]

    scr_table = _data_table(scr_df, SCR_COLS)
    tar_table = _data_table(tar_df, TAR_COLS)

    # ── Compose HTML body ──
    html = f"""
<html>
<body style="font-family:Calibri;font-size:11pt;color:#000000;margin:20px;">

<p>Dear Team,</p>

<p>
  <b>Summary (Bucketed: {buckets_str})</b> &mdash;
  SCR latest: <b>{scr_latest_str}</b>, TAR latest: <b>{tar_latest_str}</b>
</p>

{summary_html}

<br>

<p style="font-size:11pt;">
  <b>SCR</b> (Latest date in filtered set: <b>{scr_latest_str}</b>, {len(scr_df)} records)
</p>
{scr_table}

<br>

<p style="font-size:11pt;">
  <b>TAR</b> (Latest date in filtered set: <b>{tar_latest_str}</b>, {len(tar_df)} records)
</p>
{tar_table}

<br>
<p style="font-size:9pt;color:#666666;">
  Filters applied &mdash; Created Date: {target_str} &nbsp;|&nbsp;
  VSN Year: {current_year} &nbsp;|&nbsp; Days &le; 210
</p>

</body>
</html>
"""

    subject = f"Daily SCR & TAR Report \u2014 Date: {target_str}"

    send_mail(subject, html)
    print(f"✅ Mail sent for {target_date}")
    print(f"   SCR: {len(scr_df)} records | TAR: {len(tar_df)} records")


# ================= RUN =================
if __name__ == "__main__":
    main()
