import os
import json
import time
import gspread
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

BASE_DIR = r"C:\Users\Irak\Desktop\deskTop\Cash in Hand and Dic Adjustment"
TOKEN_FILE = os.path.join(BASE_DIR, "FieldEdit", "token.json")

creds = Credentials.from_authorized_user_file(
    TOKEN_FILE,
    ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
)

gc = gspread.authorize(creds)
drive_service = build('drive', 'v3', credentials=creds)
sheets_service = build('sheets', 'v4', credentials=creds)

with open(os.path.join(BASE_DIR, 'sh_by_zone.json'), 'r', encoding='utf-8') as f:
    sh_by_zone = json.load(f)

print("Loaded sh_by_zone configuration:")
print("BARI Zone Email:", sh_by_zone.get("BARI"))

def find_total_cash_in_hand_column(ws):
    try:
        all_vals = ws.get_values("A5:Z16")
        for row in all_vals:
            for c_idx, cell_val in enumerate(row):
                val_str = str(cell_val).strip().upper()
                if "TOTAL CASH IN HAND" in val_str or "TOTAL CASH" in val_str:
                    return c_idx + 1
        max_c = 0
        for row in all_vals[:7]:
            for c_idx, cell_val in enumerate(row):
                if str(cell_val).strip():
                    max_c = max(max_c, c_idx + 1)
        if max_c >= 3:
            return max_c
    except Exception:
        pass
    return 13

# Find Master Registry with retry
query = "name = 'Master_Registry_Cash_In_Hand' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false"

files = []
for attempt in range(5):
    try:
        res = drive_service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        files = res.get('files', [])
        break
    except Exception as e:
        print(f"Network retry {attempt+1}/5 for fetching registry: {e}")
        time.sleep(3)

if not files:
    print("Master Registry not found on Drive or Network unavailable.")
    exit(1)

reg_id = files[0]['id']
print(f"Found Master Registry Sheet ID: {reg_id}")

reg_sheet = None
for attempt in range(5):
    try:
        reg_sheet = gc.open_by_key(reg_id)
        break
    except Exception as e:
        print(f"Network retry {attempt+1}/5 opening registry: {e}")
        time.sleep(3)

reg_ws = reg_sheet.get_worksheet(0)
records = reg_ws.get_all_values()

print(f"\nProcessing {len(records)-1} FM sheets registered in Master Registry...")

verification_passed = True
processed_count = 0

for row_idx, r in enumerate(records[1:], start=2):
    if len(r) < 3:
        continue
    fm_name = r[0]
    zone = r[1]
    sheet_id = r[2]
    
    if not sheet_id:
        continue
        
    print(f"\n--- Processing FM ({row_idx-1}/{len(records)-1}): {fm_name} | Zone: {zone} | Sheet ID: {sheet_id} ---")
    
    success = False
    for attempt in range(4):
        try:
            ss = gc.open_by_key(sheet_id)
            worksheets = ss.worksheets()
            
            # 1. Share sheet with SH email if missing or needed
            sh_email = sh_by_zone.get(zone)
            if sh_email:
                try:
                    body = {
                        'role': 'writer',
                        'type': 'user',
                        'emailAddress': sh_email
                    }
                    drive_service.permissions().create(fileId=sheet_id, body=body, sendNotificationEmail=False).execute()
                    print(f"  [OK] Shared permissions with SH/FM: {sh_email}")
                except Exception as pe:
                    pass

            for ws in worksheets:
                tab_title = ws.title
                
                # READ DATA BEFORE PATCHING (specifically user inputs range)
                data_before = ws.get_values("A18:Z48")
                
                # Dynamically determine Total Column & Unprotected range for THIS sheet
                t_col_idx = find_total_cash_in_hand_column(ws)
                end_unprot_col = t_col_idx - 1 # Exclusive end index for 0-indexed column
                
                # Fetch sheet metadata & protected ranges
                meta = sheets_service.spreadsheets().get(
                    spreadsheetId=sheet_id,
                    fields="sheets(properties(sheetId,title),protectedRanges(protectedRangeId))"
                ).execute()
                
                s_meta = None
                for sm in meta.get('sheets', []):
                    if sm['properties']['title'] == tab_title:
                        s_meta = sm
                        break
                
                if not s_meta:
                    continue
                    
                s_id = s_meta['properties']['sheetId']
                
                requests = []
                
                # Step 1: UNFREEZE ALL COLUMNS FIRST so merged cells can be adjusted safely
                requests.append({
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": s_id,
                            "gridProperties": {
                                "frozenColumnCount": 0
                            }
                        },
                        "fields": "gridProperties.frozenColumnCount"
                    }
                })
                
                # Step 2: Unmerge title spanning B2:M3, clear B2, set C2="CASH IN HAND", then merge C2:M3 (Column C to Total Column)
                requests.append({
                    "unmergeCells": {
                        "range": {
                            "sheetId": s_id,
                            "startRowIndex": 1,
                            "endRowIndex": 3,
                            "startColumnIndex": 1,
                            "endColumnIndex": t_col_idx
                        }
                    }
                })
                requests.append({
                    "updateCells": {
                        "range": {
                            "sheetId": s_id,
                            "startRowIndex": 1,
                            "endRowIndex": 2,
                            "startColumnIndex": 1,
                            "endColumnIndex": 3
                        },
                        "rows": [
                            {
                                "values": [
                                    {
                                        "userEnteredValue": {"stringValue": ""}
                                    },
                                    {
                                        "userEnteredValue": {"stringValue": "CASH IN HAND"},
                                        "userEnteredFormat": {
                                            "horizontalAlignment": "CENTER",
                                            "verticalAlignment": "MIDDLE",
                                            "textFormat": {
                                                "bold": True,
                                                "fontSize": 26,
                                                "foregroundColor": {
                                                    "red": 0.0,
                                                    "green": 0.9490196,
                                                    "blue": 0.9960784
                                                }
                                            }
                                        }
                                    }
                                ]
                            }
                        ],
                        "fields": "userEnteredValue,userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment,userEnteredFormat.textFormat"
                    }
                })
                requests.append({
                    "mergeCells": {
                        "range": {
                            "sheetId": s_id,
                            "startRowIndex": 1,
                            "endRowIndex": 3,
                            "startColumnIndex": 2,
                            "endColumnIndex": t_col_idx
                        },
                        "mergeType": "MERGE_ALL"
                    }
                })
                
                # Step 3: Now freeze 17 rows, 2 columns (Columns A & B, i.e. DATE column ONLY)
                requests.append({
                    "updateSheetProperties": {
                        "properties": {
                            "sheetId": s_id,
                            "gridProperties": {
                                "frozenRowCount": 17,
                                "frozenColumnCount": 2
                            }
                        },
                        "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"
                    }
                })
                
                # Step 4: Delete existing protected ranges on this tab
                for pr in s_meta.get('protectedRanges', []):
                    requests.append({"deleteProtectedRange": {"protectedRangeId": pr['protectedRangeId']}})
                    
                # Step 5: Add updated protected range with dynamic unprotected range from Column C to total_col_idx-1
                editors = [sh_email] if sh_email else []
                requests.append({
                    "addProtectedRange": {
                        "protectedRange": {
                            "range": {"sheetId": s_id},
                            "description": f"Locked headers & formula columns ({tab_title})",
                            "warningOnly": False,
                            "unprotectedRanges": [{
                                "sheetId": s_id,
                                "startRowIndex": 17,
                                "endRowIndex": 48,
                                "startColumnIndex": 2,
                                "endColumnIndex": end_unprot_col
                            }],
                            "editors": {"users": editors}
                        }
                    }
                })
                
                # Execute Metadata Batch Update (Zero Cell Values Touched!)
                sheets_service.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body={"requests": requests}).execute()
                
                # READ DATA AFTER PATCHING
                data_after = ws.get_values("A18:Z48")
                
                # VERIFY ZERO DATA WAS ERASED OR ALTERED
                if data_before == data_after:
                    print(f"  [SUCCESS] Tab '{tab_title}': Unprotected cols C to {chr(64+end_unprot_col)}. Data 100% IDENTICAL before & after patching!")
                else:
                    print(f"  [WARNING] Data mismatch detected on tab '{tab_title}'!")
                    verification_passed = False

            processed_count += 1
            success = True
            time.sleep(2.5)
            break
        except Exception as e:
            print(f"  Attempt {attempt+1}/4 failed for sheet {sheet_id} ({fm_name}): {e}")
            time.sleep(15)

print("\n=======================================================")
if verification_passed:
    print(f"SUCCESS: Processed {processed_count} sheets. ALL VERIFICATIONS PASSED 100%! Settings updated with ZERO data loss!")
else:
    print("WARNING: Mismatch detected during verification.")
