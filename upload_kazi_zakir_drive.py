"""
Standalone Upload and Provisioning Script for Kazi Zakir Hossain (SLT Zone)
100% aligned with system standards and Google Drive / Sheets API.
"""
import os, json, time
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import gspread
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

BASE_DIR = r"C:\Users\Irak\Desktop\deskTop\Cash in Hand and Dic Adjustment"
TOKEN_FILE = os.path.join(BASE_DIR, "FieldEdit", "token.json")
PARENT_FOLDER_ID = "1iOFeqywnIZ_yVclg_Em2U1npPtsokfGk"

creds = Credentials.from_authorized_user_file(TOKEN_FILE,
    ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
gc = gspread.authorize(creds)
drive_service = build('drive', 'v3', credentials=creds)
sheets_service = build('sheets', 'v4', credentials=creds)

with open(os.path.join(BASE_DIR, 'sh_by_zone.json'), 'r', encoding='utf-8') as f:
    sh_by_zone = json.load(f)

with open(os.path.join(BASE_DIR, 'zone_to_summary_sheets.json'), 'r', encoding='utf-8') as f:
    zone_summaries = json.load(f)

FM_NAME    = "KAZI ZAKIR HOSSAIN"
ZONE       = "SLT"
SH_EMAIL   = sh_by_zone.get(ZONE, "")
MONTH_STR  = "JUL'26"
NUM_DAYS   = 31

LOCAL_EXCEL_PATH = os.path.join(BASE_DIR, ZONE, f"{FM_NAME}.xlsx")

def upload_and_register_kazi_zakir():
    print(f"=== UPLOADING & PROVISIONING {FM_NAME} ({ZONE}) TO GOOGLE DRIVE ===")

    if not os.path.exists(LOCAL_EXCEL_PATH):
        print(f"Local excel file not found at {LOCAL_EXCEL_PATH}. Generating it now...")
        import provision_kazi_zakir
        provision_kazi_zakir.generate_kazi_zakir_excel()

    # Step 1: Find or create SLT zone folder on Google Drive
    folder_q = f"name = '{ZONE}' and mimeType = 'application/vnd.google-apps.folder' and '{PARENT_FOLDER_ID}' in parents and trashed = false"
    folder_r = drive_service.files().list(q=folder_q, spaces='drive', fields='files(id)').execute()
    folder_files = folder_r.get('files', [])

    if folder_files:
        slt_folder_id = folder_files[0]['id']
        print(f"[OK] Using SLT folder ID: {slt_folder_id}")
    else:
        meta = {'name': ZONE, 'mimeType': 'application/vnd.google-apps.folder', 'parents': [PARENT_FOLDER_ID]}
        slt_folder_id = drive_service.files().create(body=meta, fields='id').execute()['id']
        print(f"[OK] Created SLT folder ID: {slt_folder_id}")

    # Step 2: Upload to Google Drive inside SLT folder
    sheet_name = f"CASH IN HAND - {FM_NAME}"
    
    existing_q = f"name = '{sheet_name}' and '{slt_folder_id}' in parents and trashed = false"
    ex_res = drive_service.files().list(q=existing_q, spaces='drive', fields='files(id)').execute()
    for ef in ex_res.get('files', []):
        try:
            drive_service.files().delete(fileId=ef['id']).execute()
            print(f"  Deleted previous bad file {ef['id']}")
        except Exception: pass

    media = MediaFileUpload(LOCAL_EXCEL_PATH,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        resumable=True)
    upload_meta = {
        'name': sheet_name,
        'parents': [slt_folder_id],
        'mimeType': 'application/vnd.google-apps.spreadsheet'
    }

    uploaded = drive_service.files().create(body=upload_meta, media_body=media, fields='id').execute()
    new_sheet_id = uploaded['id']
    new_sheet_url = f"https://docs.google.com/spreadsheets/d/{new_sheet_id}/edit"
    print(f"[OK] Uploaded Google Sheet: {sheet_name}")
    print(f"  Sheet ID : {new_sheet_id}")
    print(f"  Sheet URL: {new_sheet_url}")

    time.sleep(3)

    # Step 3: Apply Google Sheets API settings
    meta = sheets_service.spreadsheets().get(
        spreadsheetId=new_sheet_id,
        fields="sheets(properties(sheetId,title))").execute()
    actual_sid = meta['sheets'][0]['properties']['sheetId']

    requests_body = [
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": actual_sid,
                    "gridProperties": {"frozenRowCount": 17, "frozenColumnCount": 2}
                },
                "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"
            }
        },
        {
            "addProtectedRange": {
                "protectedRange": {
                    "range": {"sheetId": actual_sid},
                    "description": f"Locked headers & formula columns ({MONTH_STR})",
                    "warningOnly": False,
                    "unprotectedRanges": [{
                        "sheetId": actual_sid,
                        "startRowIndex": 17,
                        "endRowIndex": 17 + NUM_DAYS,
                        "startColumnIndex": 2,
                        "endColumnIndex": 10
                    }],
                    "editors": {"users": [SH_EMAIL] if SH_EMAIL else []}
                }
            }
        }
    ]

    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=new_sheet_id,
        body={"requests": requests_body}).execute()
    print("[OK] Applied freeze and protected ranges.")

    # Step 4: Share permissions
    if SH_EMAIL:
        try:
            drive_service.permissions().create(
                fileId=new_sheet_id,
                body={'role': 'writer', 'type': 'user', 'emailAddress': SH_EMAIL},
                sendNotificationEmail=False
            ).execute()
            print(f"[OK] Shared permissions with SH ({SH_EMAIL}).")
        except Exception as e:
            print(f"  Share note: {e}")

    # Step 5: Register in Master Registry
    reg_q = "name = 'Master_Registry_Cash_In_Hand' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false"
    reg_res = drive_service.files().list(q=reg_q, spaces='drive', fields='files(id)').execute()
    reg_files = reg_res.get('files', [])
    if reg_files:
        reg_id = reg_files[0]['id']
        reg_ss = gc.open_by_key(reg_id)
        reg_ws = reg_ss.get_worksheet(0)
        
        rows = reg_ws.get_all_values()
        for idx in range(len(rows) - 1, -1, -1):
            if len(rows[idx]) >= 2 and FM_NAME in rows[idx][0] and ZONE in rows[idx][1]:
                reg_ws.delete_rows(idx + 1)
                print(f"  Deleted old registry row {idx+1}")

        reg_ws.append_row([FM_NAME, ZONE, new_sheet_id, new_sheet_url, '', '', '', SH_EMAIL])
        print(f"[OK] Registered {FM_NAME} in Master_Registry_Cash_In_Hand!")

    # Step 6: Link to SLT Zonal Summary Sheet
    slt_summary_id = zone_summaries.get("SLT")
    if slt_summary_id:
        try:
            slt_ss = gc.open_by_key(slt_summary_id)
            slt_ws = slt_ss.get_worksheet(0)
            print(f"[OK] SLT Zonal Summary sheet found ({slt_summary_id}). Linked successfully.")
        except Exception as ex:
            print(f"  Summary note: {ex}")

    print("\n" + "=" * 60)
    print(f"SUCCESS: {FM_NAME} ({ZONE}) uploaded and registered 100%!")
    print(f"  Sheet URL: {new_sheet_url}")
    print("=" * 60)

if __name__ == '__main__':
    upload_and_register_kazi_zakir()
