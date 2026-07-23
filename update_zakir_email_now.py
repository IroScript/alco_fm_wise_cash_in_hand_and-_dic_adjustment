"""
Script to update Kazi Zakir Hossain's email in Master Registry and share sheet access
"""
import os, json, gspread
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

BASE_DIR = r"C:\Users\Irak\Desktop\deskTop\Cash in Hand and Dic Adjustment"
TOKEN_FILE = os.path.join(BASE_DIR, "FieldEdit", "token.json")
ZAKIR_EMAIL = "kazizakirhussain86@gmail.com"
ZAKIR_SHEET_ID = "1GushsCII0md-Ak7LewQR2GAI511oEunN_XvAg2GwJpU"

def update_email():
    if not os.path.exists(TOKEN_FILE):
        print("token.json not found.")
        return

    creds = Credentials.from_authorized_user_file(
        TOKEN_FILE,
        ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    )
    gc = gspread.authorize(creds)
    drive_service = build('drive', 'v3', credentials=creds)

    print("1. Updating Master_Registry_Cash_In_Hand with kazizakirhussain86@gmail.com...")
    reg_q = "name = 'Master_Registry_Cash_In_Hand' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false"
    reg_res = drive_service.files().list(q=reg_q, spaces='drive', fields='files(id)').execute()
    reg_files = reg_res.get('files', [])

    if reg_files:
        reg_id = reg_files[0]['id']
        reg_ss = gc.open_by_key(reg_id)
        reg_ws = reg_ss.get_worksheet(0)
        rows = reg_ws.get_all_values()
        for idx, row in enumerate(rows):
            if len(row) > 0 and 'KAZI ZAKIR HOSSAIN' in str(row[0]).upper():
                reg_ws.update_cell(idx + 1, 5, ZAKIR_EMAIL)
                print(f"[OK] Updated Master Registry row {idx+1} FM Email to {ZAKIR_EMAIL}")

    print("2. Sharing Kazi Zakir Sheet directly with kazizakirhussain86@gmail.com...")
    try:
        drive_service.permissions().create(
            fileId=ZAKIR_SHEET_ID,
            body={'role': 'writer', 'type': 'user', 'emailAddress': ZAKIR_EMAIL},
            sendNotificationEmail=False
        ).execute()
        print(f"[OK] Shared writer permissions with {ZAKIR_EMAIL}")
    except Exception as e:
        print("Share exception:", e)

    print("\n=== EMAIL UPDATE & PERMISSION SHARE COMPLETE ===")

if __name__ == '__main__':
    update_email()
