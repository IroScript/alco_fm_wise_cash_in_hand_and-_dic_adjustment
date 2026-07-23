"""
Script to check all files in MYM.A folder and remove non-Cash-In-Hand files (DPSR, etc.)
"""
import os, json, gspread
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

BASE_DIR = r"C:\Users\Irak\Desktop\deskTop\Cash in Hand and Dic Adjustment"
TOKEN_FILE = os.path.join(BASE_DIR, "FieldEdit", "token.json")
MYM_A_FOLDER_ID = "1BAURICLLBLs1AAi2B_IeX-vg3R9oAQ4N"

def clean_mym_a():
    if not os.path.exists(TOKEN_FILE):
        print("token.json not found.")
        return

    creds = Credentials.from_authorized_user_file(
        TOKEN_FILE,
        ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    )
    ds = build('drive', 'v3', credentials=creds)

    print("Checking all files inside MYM.A folder...")
    res = ds.files().list(
        q=f"'{MYM_A_FOLDER_ID}' in parents and trashed = false",
        fields="files(id, name, parents)"
    ).execute()
    files = res.get('files', [])

    print(f"Total files in MYM.A: {len(files)}")
    for f in files:
        fname = f['name']
        fid = f['id']
        print(f"File: {fname} (ID: {fid})")

        # If file is DPSR or not CASH IN HAND, remove MYM.A parent from it
        if "DPSR" in fname.upper() or "JOB" in fname.upper() or "TP" in fname.upper() or not fname.upper().startswith("CASH IN HAND"):
            print(f" -> Removing non-Cash-In-Hand file '{fname}' from MYM.A folder...")
            ds.files().update(
                fileId=fid,
                removeParents=MYM_A_FOLDER_ID,
                fields="id, parents"
            ).execute()
            print(f" [OK] Removed '{fname}' from MYM.A!")

    print("\n=== CLEANUP OF MYM.A FOLDER COMPLETE ===")

if __name__ == '__main__':
    clean_mym_a()
