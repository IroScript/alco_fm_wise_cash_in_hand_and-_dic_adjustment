"""
Script to update SLT Zonal Summary Google Sheet with Kazi Zakir Hossain's sheet IMPORTRANGE formulas
"""
import os, json, time
import gspread
from google.oauth2.credentials import Credentials

BASE_DIR = r"C:\Users\Irak\Desktop\deskTop\Cash in Hand and Dic Adjustment"
TOKEN_FILE = os.path.join(BASE_DIR, "FieldEdit", "token.json")
SLT_SUMMARY_ID = "1fWlppljj12hQZjZaPJmCBlXTIx4WI3y0bNL2zXVqtE0"
KAZI_ZAKIR_SHEET_URL = "https://docs.google.com/spreadsheets/d/1GushsCII0md-Ak7LewQR2GAI511oEunN_XvAg2GwJpU/edit"

NUM_DAYS = 31

def update_slt_summary():
    if not os.path.exists(TOKEN_FILE):
        print("token.json not found.")
        return

    creds = Credentials.from_authorized_user_file(
        TOKEN_FILE,
        ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    )
    gc = gspread.authorize(creds)
    ss = gc.open_by_key(SLT_SUMMARY_ID)
    ws = ss.get_worksheet(0)

    print("Connected to SLT Zonal Summary sheet...")

    # Update Row 6 cell F6 to KAZI ZAKIR HOSSAIN
    ws.update_cell(6, 6, "KAZI ZAKIR HOSSAIN")
    print("[OK] Updated Row 6 Header for KAZI ZAKIR HOSSAIN.")

    # Update Row 7 cells Q7:V7 to KAZI ZAKIR HOSSAIN
    ws.update("Q7:V7", [["KAZI ZAKIR HOSSAIN"] * 6])
    print("[OK] Updated Row 7 FM labels for KAZI ZAKIR HOSSAIN.")

    cell_updates = []

    # Update column F row 18:48 formula for Kazi Zakir TOTAL CASH IN HAND
    for idx in range(NUM_DAYS):
        r = 18 + idx
        formula_total = f'=IMPORTRANGE("{KAZI_ZAKIR_SHEET_URL}", "JUL\'26!J{r}")'
        cell_updates.append({'range': f'F{r}', 'values': [[formula_total]]})

        # Detail IMPORTRANGE for FM self + MPOs (cols Q through V)
        formula_details = f'=IMPORTRANGE("{KAZI_ZAKIR_SHEET_URL}", "JUL\'26!C{r}:H{r}")'
        cell_updates.append({'range': f'Q{r}', 'values': [[formula_details]]})

    ws.batch_update(cell_updates, value_input_option='USER_ENTERED')
    print("[OK] Added IMPORTRANGE formulas to SLT Zonal Summary for Kazi Zakir Hossain!")

if __name__ == '__main__':
    update_slt_summary()
