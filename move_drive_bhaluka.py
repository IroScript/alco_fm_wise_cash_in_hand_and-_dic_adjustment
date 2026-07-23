"""
Script to move Bhaluka Sheet to MYM.A folder on Google Drive and update Master Registry
for account: store.alcopharma@gmail.com
"""
import os, json, time
import gspread
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

BASE_DIR = r"C:\Users\Irak\Desktop\deskTop\Cash in Hand and Dic Adjustment"
CLIENT_SECRET_FILE = os.path.join(BASE_DIR, "FieldEdit", "client_secret_866102064521-5g6tq5989nqs97ehgse7n6fl1o9pslt5.apps.googleusercontent.com.json")
TOKEN_FILE = os.path.join(BASE_DIR, "FieldEdit", "token.json")

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

MYM_A_FOLDER_ID = "1BAURICLLBLs1AAi2B_IeX-vg3R9oAQ4N"

def get_creds():
    creds = None
    if os.path.exists(TOKEN_FILE):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        except Exception:
            creds = None
    if not creds or not creds.valid:
        refreshed = False
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                refreshed = True
            except Exception as e:
                print(f"Token refresh error: {e}")
                creds = None
        if not refreshed:
            if os.path.exists(TOKEN_FILE):
                try: os.remove(TOKEN_FILE)
                except Exception: pass
            print("Opening browser for Google authorization (Please select store.alcopharma@gmail.com)...")
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            creds = flow.run_local_server(port=0, prompt='select_account')
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    return creds

def move_bhaluka_on_drive():
    print("=== MOVING BHALUKA SHEET TO MYM.A FOLDER ON GOOGLE DRIVE ===")
    creds = get_creds()
    gc = gspread.authorize(creds)
    drive_service = build('drive', 'v3', credentials=creds)

    # 1. Find Bhaluka Sheet on Google Drive
    query = "(name contains 'BHALUKA' or name contains 'Bhaluka') and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false"
    res = drive_service.files().list(q=query, spaces='drive', fields='files(id, name, parents)').execute()
    files = res.get('files', [])

    if not files:
        print("⚠ Could not find Bhaluka sheet on Google Drive.")
    else:
        for f in files:
            file_id = f['id']
            file_name = f['name']
            curr_parents = f.get('parents', [])
            print(f"Found sheet: {file_name} (ID: {file_id}, Current Parents: {curr_parents})")

            # Move to MYM.A folder
            previous_parents = ",".join(curr_parents)
            updated_file = drive_service.files().update(
                fileId=file_id,
                addParents=MYM_A_FOLDER_ID,
                removeParents=previous_parents,
                fields='id, parents'
            ).execute()
            print(f"✔ Successfully moved '{file_name}' to MYM.A folder ({MYM_A_FOLDER_ID})!")

    # 2. Update Master Registry Cash In Hand
    reg_q = "name = 'Master_Registry_Cash_In_Hand' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false"
    reg_res = drive_service.files().list(q=reg_q, spaces='drive', fields='files(id)').execute()
    reg_files = reg_res.get('files', [])
    if reg_files:
        reg_id = reg_files[0]['id']
        reg_ss = gc.open_by_key(reg_id)
        reg_ws = reg_ss.get_worksheet(0)
        rows = reg_ws.get_all_values()
        for idx, row in enumerate(rows):
            if len(row) > 1 and ('BHALUKA' in str(row[0]).upper() or 'BHALUKA' in str(row[2]).upper()):
                reg_ws.update_cell(idx + 1, 2, 'MYM.A')
                print(f"✔ Updated Master Registry row {idx+1} zone to MYM.A for {row[0]}")

if __name__ == '__main__':
    move_bhaluka_on_drive()
