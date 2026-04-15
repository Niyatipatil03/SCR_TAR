# =========================================================
# DAILY SCR & TAR MAIL SCRIPT
# Sends ONE mail per day
# Filters:
#   - Created date = target day
#   - VSN Year = 2026
#   - Days <= 210
# =========================================================

import os
import pandas as pd
import win32com.client as win32
from datetime import date, timedelta

# ================= CONFIG =================
COMPILED_FILE = (
    r"C:\Users\25033810\OneDrive - Mahindra & Mahindra Ltd"
    r"\Field Quality Nasik - SCR & TAR\SCR_TAR_Compiled.xlsx"
)

MAIL_TO = "Patil.Niyati@mahindra.com"
MAIL_CC = ""

DATE_PARSED_COL = "Created date(parsed)"
VSN_DATE_COL = "VSN Date"
DAYS_COL = "Days"
REPORT_COL = "Report No."
MODEL_COL = "Model Family"
VEHICLE_COL = "Vehicle Sr No"
DESC_COL = "Description"

# ================= HELPERS =================
def detect_type(r):
    r = str(r).upper()
    if r.startswith("S-"):
        return "SCR"
    if r.startswith("T-"):
        return "TAR"
    return "UNKNOWN"

def send_mail(subject, html):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.To = MAIL_TO
    mail.CC = MAIL_CC
    mail.Subject = subject
    mail.HTMLBody = html
    mail.Send()

# ================= MAIN =================
def main():

    # ✅ Safe Excel read
    try:
        df = pd.read_excel(COMPILED_FILE, engine="openpyxl")
    except PermissionError:
        raise RuntimeError(
            "❌ SCR_TAR_Compiled.xlsx is open or locked.\n"
            "Please close Excel and wait for OneDrive sync, then rerun."
        )

    # ✅ Normalize column names
    df.columns = (
        df.columns.astype(str)
        .str.replace("\u00A0", " ", regex=False)
        .str.strip()
    )

    # ✅ Ensure Created date(parsed)
    if DATE_PARSED_COL not in df.columns:
        df[DATE_PARSED_COL] = pd.to_datetime(
            df["Created date"],
            errors="coerce",
            dayfirst=True,
            infer_datetime_format=True
        ).dt.date

    # ✅ Convert fixed columns once
    df[DATE_PARSED_COL] = pd.to_datetime(df[DATE_PARSED_COL], errors="coerce").dt.date
    df[VSN_DATE_COL] = pd.to_datetime(df[VSN_DATE_COL], errors="coerce")
    df[DAYS_COL] = pd.to_numeric(df[DAYS_COL], errors="coerce")

    # ================= TARGET DATES =================
    dates_to_send = [
        date.today() - timedelta(days=1),  # Yesterday (14 Apr)
        date.today() - timedelta(days=2),  # Day before (13 Apr)
    ]

    for target_date in dates_to_send:

        # ✅ Apply filters
        filtered = df[
            (df[DATE_PARSED_COL] == target_date) &
            (df[DAYS_COL] <= 210) &
            (df[VSN_DATE_COL].dt.year == 2026)
        ]

        if filtered.empty:
            print(f"ℹ️ No data for {target_date}. Mail not sent.")
            continue

        # SCR / TAR split
        filtered["Type"] = filtered[REPORT_COL].apply(detect_type)
        scr_df = filtered[filtered["Type"] == "SCR"]
        tar_df = filtered[filtered["Type"] == "TAR"]

        # ================= MAIL BODY =================
        html = f"""
        <html>
        <body style="font-family:Calibri; font-size:11pt">

            <p>Dear Team,</p>

            <p>
                <b>Filter Applied</b><br>
                • Created Date: <b>{target_date.strftime('%d-%b-%Y')}</b><br>
                • VSN Year: <b>2026</b><br>
                • Days ≤ <b>210</b>
            </p>

            <h3>SCR ({len(scr_df)})</h3>
            {scr_df[[REPORT_COL, MODEL_COL, VEHICLE_COL, DESC_COL, DAYS_COL]]
                .to_html(index=False, border=0)}

            <h3>TAR ({len(tar_df)})</h3>
            {tar_df[[REPORT_COL, MODEL_COL, VEHICLE_COL, DESC_COL, DAYS_COL]]
                .to_html(index=False, border=0)}

        </body>
        </html>
        """

        subject = (
            f"Daily SCR & TAR Report | "
            f"{target_date.strftime('%d-%b-%Y')} | "
            f"VSN 2026 | Days ≤ 210"
        )

        send_mail(subject, html)
        print(f"✅ Mail sent successfully for {target_date}")

# ================= RUN =================
if __name__ == "__main__":
    main()