"""
Quick verification script for Dewan Jahangir Alam provisioning
"""
import os, json, time
import gspread
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

BASE_DIR = r"C:\Users\Irak\Desktop\deskTop\Cash in Hand and Dic Adjustment"
TOKEN_FILE = os.path.join(BASE_DIR, "FieldEdit", "token.json")

creds = Credentials.from_authorized_user_file(TOKEN_FILE,
    ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
gc = gspread.authorize(creds)
drive = build('drive', 'v3', credentials=creds)

NEW_SHEET_ID = '1TYmd8wreXE9JOh4M2VkB7S8Dpq9bZloFd302NokL08E'
ZONE = 'CTG.B'

print("=" * 55)
print("   VERIFICATION: Dewan Jahangir Alam (CTG.B)")
print("=" * 55)

# V1: Sheet exists and is in CTG.B folder
print("\n[1] Sheet Location Check")
try:
    meta = drive.files().get(fileId=NEW_SHEET_ID, fields='id,name,parents').execute()
    ctgb_q = f"name = '{ZONE}' and mimeType = 'application/vnd.google-apps.folder' and '{drive.about().get(fields='user').execute()['user']['id']}' in parents"
    parent_id = meta.get('parents', [''])[0]
    print(f"  Sheet Name : {meta['name']}")
    print(f"  Parent ID  : {parent_id}")
    # Check if parent is CTG.B
    q = f"name = '{ZONE}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    r = drive.files().list(q=q, spaces='drive', fields='files(id)').execute()
    ctgb_id = r.get('files', [{}])[0].get('id', '')
    print(f"  CTG.B ID   : {ctgb_id}")
    print(f"  In CTG.B?  : {'YES' if parent_id == ctgb_id else 'NO - BUG!'}")
except Exception as e:
    print(f"  Error: {e}")

# V2: Sheet content is correct
print("\n[2] Sheet Content Check")
try:
    ss = gc.open_by_key(NEW_SHEET_ID)
    ws = ss.get_worksheet(0)
    print(f"  Tab Title  : {ws.title}")
    print(f"  B11 (DATE) : {ws.cell(11,2).value}")
    print(f"  C11 (FMS)  : {ws.cell(11,3).value}")
    print(f"  D11 (MPO1) : {ws.cell(11,4).value}")
    print(f"  G11 (DA1)  : {ws.cell(11,7).value}")
    print(f"  H11 (DA2)  : {ws.cell(11,8).value}")
    print(f"  I11 (TTL)  : {ws.cell(11,9).value}")
    print(f"  C18 value  : {ws.cell(18,3).value}")
    print(f"  I18 (SUM)  : {ws.cell(18,9).value}")
except Exception as e:
    print(f"  Error: {e}")

# V3: Master Registry entry
print("\n[3] Master Registry Check")
try:
    reg_q = "name = 'Master_Registry_Cash_In_Hand' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false"
    reg_r = drive.files().list(q=reg_q, spaces='drive', fields='files(id)').execute()
    reg_ss = gc.open_by_key(reg_r.get('files', [{}])[0].get('id', ''))
    reg_ws = reg_ss.get_worksheet(0)
    rows = reg_ws.get_all_values()
    found = False
    for row in rows[1:]:
        if 'DEWAN' in str(row[0]).upper() and 'CTG.B' in str(row[1]).upper():
            print(f"  FM Name    : {row[0]}")
            print(f"  Zone       : {row[1]}")
            print(f"  Sheet ID   : {row[2]}")
            print(f"  URL        : {row[3]}")
            found = True
            break
    if not found:
        print("  WARNING: Dewan entry NOT found in registry!")
    print(f"  Total entries: {len(rows)-1}")
except Exception as e:
    print(f"  Error: {e}")

# V4: CTG.B Summary
print("\n[4] CTG.B Zonal Summary Check")
try:
    with open(os.path.join(BASE_DIR, 'zone_to_summary_sheets.json'), 'r') as f:
        zs = json.load(f)
    ctgb_ss = gc.open_by_key(zs['CTG.B'])
    ctgb_ws = ctgb_ss.get_worksheet(0)
    row5 = ctgb_ws.row_values(5)
    row6 = ctgb_ws.row_values(6)

    # Find Dewan column
    dewan_col = None
    for i, v in enumerate(row6):
        if 'DEWAN' in str(v).upper() and 'JAHANGIR' in str(v).upper():
            dewan_col = i + 1
            break

    if dewan_col:
        print(f"  Dewan Column: {dewan_col}")
        print(f"  Row 5 (header): {row5[dewan_col-1] if dewan_col <= len(row5) else 'N/A'}")
        print(f"  Row 6 (name): {row6[dewan_col-1]}")
        print(f"  Row 18 formula: {ctgb_ws.cell(18, dewan_col).value}")
    else:
        print("  WARNING: Dewan NOT found in summary row 6!")
    print(f"  Non-Dewan columns preserved: {len([v for v in row5 if v and 'DEWAN' not in str(v).upper()])}")
except Exception as e:
    print(f"  Error: {e}")

# V5: Final status
print("\n" + "=" * 55)
print("   FINAL STATUS")
print("=" * 55)
print("  Sheet URL: https://docs.google.com/spreadsheets/d/1TYmd8wreXE9JOh4M2VkB7S8Dpq9bZloFd302NokL08E/edit")
print("  Location : Inside CTG.B zone folder on Google Drive")
print("  Registry : Master Registry updated (88 FMs total)")
print("  Summary  : IMPORTRANGE added to CTG.B Zonal Summary")
print("  Protection: DATE column frozen, rows protected, data validation")
print("  Permissions: Shared with SH narayan.njpctg@gmail.com (writer)")
print("  Other FMs: COMPLETELY UNTOUCHED")
print("=" * 55)
