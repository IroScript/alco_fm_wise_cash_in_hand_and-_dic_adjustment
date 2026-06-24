import os
import re
import time
import argparse
import datetime
import calendar
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

BASE_DIR = r"c:\Users\Irak\Desktop\Cash in Hand and Dic Adjustment"
CREDS_PATH = os.path.join(BASE_DIR, "FieldEdit", "alco-pharma-cf4b49e394bb.json")
PARENT_FOLDER_ID = "1iOFeqywnIZ_yVclg_Em2U1npPtsokfGk"

# Setup Google APIs
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = Credentials.from_service_account_file(CREDS_PATH, scopes=SCOPES)
gc = gspread.authorize(creds)
drive_service = build('drive', 'v3', credentials=creds)
sheets_service = build('sheets', 'v4', credentials=creds)

def get_col_letter(col_idx):
    # Standard 1-based index to A, B, C...
    result = ""
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result = chr(65 + remainder) + result
    return result

def get_month_info(target_date):
    # e.g., target_date = 2026-07-01
    # returns:
    # current_month_str: "JUL'26"
    # prev_month_str: "JUN'26"
    # num_days_current: 31
    # num_days_prev: 30
    # current_month_short: "JUL"
    # year_short: "26"
    
    year = target_date.year
    month = target_date.month
    
    current_month_short = calendar.month_abbr[month].upper()
    year_short = str(year)[2:]
    current_month_str = f"{current_month_short}'{year_short}"
    
    # Calculate previous month
    if month == 1:
        prev_month = 12
        prev_year = year - 1
    else:
        prev_month = month - 1
        prev_year = year
        
    prev_month_short = calendar.month_abbr[prev_month].upper()
    prev_year_short = str(prev_year)[2:]
    prev_month_str = f"{prev_month_short}'{prev_year_short}"
    
    _, num_days_current = calendar.monthrange(year, month)
    _, num_days_prev = calendar.monthrange(prev_year, prev_month)
    
    return current_month_str, prev_month_str, num_days_current, num_days_prev

def protect_sheet(spreadsheet_id, sheet_id, user_email, message="Locked archive"):
    try:
        body = {
            "requests": [
                {
                    "addProtectedRange": {
                        "protectedRange": {
                            "range": {
                                "sheetId": sheet_id
                            },
                            "description": message,
                            "warningOnly": False,
                            "editors": {
                                "users": [user_email]
                            }
                        }
                    }
                }
            ]
        }
        sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body).execute()
        print(f"Locked sheet ID: {sheet_id} successfully.")
    except Exception as e:
        print(f"Error locking sheet ID {sheet_id}: {e}")

def rollover_fm_sheet(sheet_id, fm_name, target_date, user_email):
    current_m, prev_m, curr_days, prev_days = get_month_info(target_date)
    print(f"Rolling over sheet {sheet_id} for FM: {fm_name} | {prev_m} -> {current_m}")
    
    try:
        ss = gc.open_by_key(sheet_id)
        
        # Check if new month tab already exists
        for ws in ss.worksheets():
            if ws.title == current_m:
                print(f"New month sheet {current_m} already exists. Skipping.")
                return
                
        # Get previous month's sheet
        try:
            prev_ws = ss.worksheet(prev_m)
        except gspread.WorksheetNotFound:
            # Fallback to the first worksheet if prev month tab name is not found
            prev_ws = ss.get_worksheet(0)
            print(f"Warning: Sheet '{prev_m}' not found. Using '{prev_ws.title}' as template.")
            
        # Duplicate the template sheet to create the new month sheet at index 0
        new_ws = ss.duplicate_sheet(source_sheet_id=prev_ws.id, insert_sheet_index=0, new_sheet_name=current_m)
        print(f"Duplicated '{prev_ws.title}' as '{current_m}' at index 0.")
        
        # Calculate number of columns
        total_col = len(new_ws.row_values(11))
        total_col_letter = get_col_letter(total_col)
        
        # Lock/Protect the old sheet
        protect_sheet(sheet_id, prev_ws.id, user_email, message=f"Locked archive for {prev_m}")
        
        # Modify the new sheet to match the current month's dates and clear data
        # Check row count for date data in duplicated sheet. It had prev_days.
        # July dates start at row 18.
        # We need curr_days of date rows.
        diff = curr_days - prev_days
        
        if diff > 0:
            # Insert rows before the last date row (row 18 + prev_days - 1)
            insert_row_idx = 18 + prev_days
            body = {
                "requests": [
                    {
                        "insertDimension": {
                            "range": {
                                "sheetId": new_ws.id,
                                "dimension": "ROWS",
                                "startIndex": insert_row_idx - 1, # 0-indexed
                                "endIndex": insert_row_idx - 1 + diff
                            },
                            "inheritFromBefore": True
                        }
                    }
                ]
            }
            sheets_service.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body=body).execute()
            print(f"Inserted {diff} rows in sheet.")
        elif diff < 0:
            # Delete rows
            delete_row_idx = 18 + curr_days
            body = {
                "requests": [
                    {
                        "deleteDimension": {
                            "range": {
                                "sheetId": new_ws.id,
                                "dimension": "ROWS",
                                "startIndex": delete_row_idx - 1,
                                "endIndex": delete_row_idx - 1 - diff
                            }
                        }
                    }
                ]
            }
            sheets_service.spreadsheets().batchUpdate(spreadsheetId=sheet_id, body=body).execute()
            print(f"Deleted {abs(diff)} rows from sheet.")
            
        # Update dates in Column B and clear inputs in Columns C to total_col - 1
        # Also re-write SUM formulas in the total column
        date_vals = [f"{d} {current_m}" for d in range(curr_days, 0, -1)]
        
        # Batch update values for speed
        cell_updates = []
        
        # 1. Update dates
        for idx, d_val in enumerate(date_vals):
            r_num = 18 + idx
            cell_updates.append({
                'range': f'B{r_num}',
                'values': [[d_val]]
            })
            
        # 2. Clear input cells and set formulas for the new rows
        # Clears range C18: [LastInputCol][18 + curr_days - 1]
        last_input_col_letter = get_col_letter(total_col - 1)
        clear_range = f"C18:{last_input_col_letter}{18 + curr_days - 1}"
        new_ws.batch_clear([clear_range])
        print("Cleared data input cells.")
        
        # 3. Write SUM formulas for all rows
        for idx in range(curr_days):
            r_num = 18 + idx
            sum_formula = f"=SUM(C{r_num}:{last_input_col_letter}{r_num})"
            cell_updates.append({
                'range': f'{total_col_letter}{r_num}',
                'values': [[sum_formula]]
            })
            
        new_ws.batch_update(cell_updates)
        print("Updated dates and formulas successfully.")
        
    except Exception as e:
        print(f"Error rolling over FM sheet {sheet_id}: {e}")

def update_boss_summary_sheets(registry, target_date):
    current_m, prev_m, curr_days, prev_days = get_month_info(target_date)
    print(f"\n--- Updating Boss Summary Sheets for {current_m} ---")
    
    # Group registry by Boss Name / Boss Email
    boss_groups = {}
    for r in registry:
        boss_name = r.get('Boss Name')
        boss_email = r.get('Boss Email')
        fm_name = r.get('FM Name')
        sheet_url = r.get('URL')
        
        if not boss_name or "VACANT" in boss_name.upper():
            continue
            
        key = (boss_name, boss_email)
        if key not in boss_groups:
            boss_groups[key] = []
        boss_groups[key].append({
            'fm_name': fm_name,
            'sheet_url': sheet_url
        })
        
    for (boss_name, boss_email), fms in boss_groups.items():
        summary_name = f"CASH IN HAND SUMMARY - {boss_name}"
        print(f"Creating/Updating summary sheet for Boss: {boss_name} ({boss_email})")
        
        # Search if summary sheet already exists
        query = f"name = '{summary_name}' and mimeType = 'application/vnd.google-apps.spreadsheet' and '{PARENT_FOLDER_ID}' in parents and trashed = false"
        res = drive_service.files().list(q=query, spaces='drive', fields='files(id)').execute()
        files = res.get('files', [])
        
        if files:
            boss_sheet_id = files[0]['id']
            boss_ss = gc.open_by_key(boss_sheet_id)
        else:
            file_metadata = {
                'name': summary_name,
                'mimeType': 'application/vnd.google-apps.spreadsheet',
                'parents': [PARENT_FOLDER_ID]
            }
            boss_file = drive_service.files().create(body=file_metadata, fields='id').execute()
            boss_sheet_id = boss_file['id']
            boss_ss = gc.open_by_key(boss_sheet_id)
            print(f"Created new Boss Summary sheet: {summary_name}")
            
            # Share with Boss
            if boss_email:
                try:
                    user_permission = {
                        'type': 'user',
                        'role': 'writer',
                        'emailAddress': boss_email
                    }
                    drive_service.permissions().create(fileId=boss_sheet_id, body=user_permission).execute()
                    print(f"Shared boss summary sheet with {boss_email}")
                except Exception as e:
                    print(f"Error sharing summary sheet with {boss_email}: {e}")
                    
        # Check if tab for current_m exists in summary sheet
        ws_exists = False
        for ws in boss_ss.worksheets():
            if ws.title == current_m:
                ws_exists = True
                boss_ws = ws
                break
                
        if not ws_exists:
            boss_ws = boss_ss.add_worksheet(title=current_m, rows=100, cols=20)
            print(f"Added tab {current_m} to Boss Summary sheet.")
        else:
            boss_ws.clear()
            print(f"Cleared existing tab {current_m} in Boss Summary sheet.")
            
        # Write IMPORTRANGE headers and values
        # We will write:
        # Col A: FM Name
        # Col B: Real-time IMPORTRANGE formula pointing to the sum cell of each date in the FM sheet
        # For simplicity, let's list dates horizontally or vertically.
        # Standard design: Let's list dates vertically (Column A) and FMs horizontally (Column B, C, D...)
        # Row 1: Header (FM Name)
        # Row 2: Dates and IMPORTRANGES
        
        headers = ["DATE"] + [f["fm_name"] for f in fms]
        rows = [headers]
        
        # July dates
        date_rows = [f"{d} {current_m}" for d in range(curr_days, 0, -1)]
        
        for idx, date_val in enumerate(date_rows):
            row_data = [date_val]
            for fm_idx, fm_info in enumerate(fms):
                fm_sheet_url = fm_info['sheet_url']
                # Determine the total column letter of that FM sheet.
                # Since we don't want to load every sheet to check its columns, we can check how many columns it has or just write a generic importrange referencing the total column.
                # Actually, let's use the formula IMPORTRANGE. The total column index is total_col.
                # Let's read the number of columns in the first sheet or query the columns.
                # Alternatively, we can use the formula pointing to the TOTAL column.
                # In our design, total_col = 4 + num_mpos + num_das.
                # Let's query the sheet columns from the FM sheet or use a formula that dynamically finds the last column.
                # But Sheets API is fast if we just reference the specific column index. Let's lookup FM sheet total col letter.
                # Since we have the registry, let's query the spreadsheet.
                # To be fast, let's just fetch the sheet properties of the FM sheet once.
                try:
                    fm_ss = gc.open_by_key(fm_sheet_url.split('/d/')[1].split('/')[0])
                    fm_ws = fm_ss.worksheet(current_m)
                    total_cols_count = len(fm_ws.row_values(11))
                    total_col_let = get_col_letter(total_cols_count)
                    
                    # We want to import the values from Row 18 down to 18 + curr_days - 1
                    # Formula: =IMPORTRANGE("URL", "JUL'26!ColumnLetterRow")
                    # For cell B18, B19, etc.
                    # Or we can do one IMPORTRANGE for the whole column:
                    # =IMPORTRANGE("URL", "JUL'26!K18:K48")
                    # This imports the whole column vertically! This is extremely elegant!
                    # So we just write the IMPORTRANGE in row 2 of each FM column.
                except Exception as e:
                    print(f"Error querying column count for FM sheet: {e}")
                    total_col_let = "K" # fallback
                    
                import_range_formula = f'=IMPORTRANGE("{fm_sheet_url}", "{current_m}!{total_col_let}18:{total_col_let}{18 + curr_days - 1}")'
                row_data.append(import_range_formula)
                
            rows.append(row_data)
            break # We only need to write the IMPORTRANGE formulas once in Row 2, and Google Sheets will spill them down vertically!
            
        # We need to write the date column vertically for the rest of the rows
        # Since IMPORTRANGE spills down, we can populate Column A with the dates manually so they align perfectly.
        # Rows list:
        # Row 1: Headers (DATE, FM1, FM2, ...)
        # Row 2: [31 JUL'26, =IMPORTRANGE(FM1), =IMPORTRANGE(FM2), ...]
        # Row 3: [30 JUL'26]
        # Row 4: [29 JUL'26]
        # ...
        full_rows = [headers]
        # Row 2 contains date and formulas
        row2 = [date_rows[0]]
        for fm_info in fms:
            try:
                fm_ss = gc.open_by_key(fm_info['sheet_url'].split('/d/')[1].split('/')[0])
                fm_ws = fm_ss.worksheet(current_m)
                total_cols_count = len(fm_ws.row_values(11))
                total_col_let = get_col_letter(total_cols_count)
            except Exception as e:
                total_col_let = "K"
            formula = f'=IMPORTRANGE("{fm_info["sheet_url"]}", "{current_m}!{total_col_let}18:{total_col_let}{18 + curr_days - 1}")'
            row2.append(formula)
        full_rows.append(row2)
        
        # Remaining rows only contain the date in Col A
        for d_val in date_rows[1:]:
            full_rows.append([d_val])
            
        boss_ws.update('A1', full_rows, value_input_option='USER_ENTERED')
        print(f"Boss summary sheet updated for {boss_name}")

def main():
    parser = argparse.ArgumentParser(description="Google Sheets Monthly Rollover Automation")
    parser.add_argument("--simulate-date", help="Simulation target date (YYYY-MM-DD)", default=None)
    args = parser.parse_args()
    
    if args.simulate_date:
        target_date = datetime.datetime.strptime(args.simulate_date, "%Y-%m-%d").date()
        print(f"SIMULATION RUN: Active Date set to {target_date}")
    else:
        target_date = datetime.date.today()
        print(f"NORMAL RUN: Active Date set to {target_date}")
        
    # Read Master Registry Sheet
    print("Reading Master Registry...")
    user_email = drive_service.about().get(fields="user(emailAddress)").execute()['user']['emailAddress']
    # Search for registry
    registry_name = "Master_Registry_Cash_In_Hand"
    query = f"name = '{registry_name}' and mimeType = 'application/vnd.google-apps.spreadsheet' and '{PARENT_FOLDER_ID}' in parents and trashed = false"
    res = drive_service.files().list(q=query, spaces='drive', fields='files(id)').execute()
    files = res.get('files', [])
    if not files:
        print("Error: Master Registry sheet not found in Drive!")
        return
        
    reg_sheet = gc.open_by_key(files[0]['id'])
    reg_ws = reg_sheet.get_worksheet(0)
    rows = reg_ws.get_all_records()
    
    print(f"Found {len(rows)} FMs in registry.")
    
    for r in rows:
        sheet_id = r.get('Sheet ID')
        fm_name = r.get('FM Name')
        if sheet_id and fm_name:
            rollover_fm_sheet(sheet_id, fm_name, target_date, user_email)
            time.sleep(1)
            
    # Update Boss Summary Sheets
    update_boss_summary_sheets(rows, target_date)
    print("Monthly rollover run completed.")

if __name__ == "__main__":
    main()
