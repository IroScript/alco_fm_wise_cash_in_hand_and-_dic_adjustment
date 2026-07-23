"""
Script to restore accidentally moved non-Cash-In-Hand files back to their original Google Drive folders
"""
import os, json, gspread
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

BASE_DIR = r"C:\Users\Irak\Desktop\deskTop\Cash in Hand and Dic Adjustment"
TOKEN_FILE = os.path.join(BASE_DIR, "FieldEdit", "token.json")
MYM_A_FOLDER_ID = "1BAURICLLBLs1AAi2B_IeX-vg3R9oAQ4N"

RESTORE_MAPPING = [
    {"id": "10sIS-KK8MyWGFs5zY06Wqgqi35vaCxKl", "orig_parent": "1pVGWzOKkyJEKzBabW4A_uOMVFMXDHMOb", "name": "DPSR, VACANT, BHALUKA"},
    {"id": "1HvCmPaOLormggvw7O6gsJWCM9acwNQMC", "orig_parent": "1ZI8PYw-RSZtyLP8zWAmDG_onTRtNrvbQ", "name": "JOB.S, VACANT, BHALUKA, MYM.B"},
    {"id": "1woLxhVhGpk5waPCwv_7IBJZFlnmbuFxT", "orig_parent": "1RxsvOj6GMgKBU1LOw3lARJijG84amt9Q", "name": "JOB.F, VACANT, BHALUKA, MYM.B"},
    {"id": "18pH55R4DQal4KYTb7oF50ZeAMXLViGx4", "orig_parent": "1Q7cLf4BnicwEJd9FTSIGGFzqhOln2cM3", "name": "TP, BHALUKA"},
    {"id": "1TXA6FIn3PDxNLhaJ4Dc-4T7c1Qr3bXx2", "orig_parent": "1HvCmPaOLormggvw7O6gsJWCM9acwNQMC", "name": "JOB.S,  BHALUKA"},
    {"id": "1FaMcX3tkwknXGrOo5e7IZxUSfAXpMdHp", "orig_parent": "1woLxhVhGpk5waPCwv_7IBJZFlnmbuFxT", "name": "JOB.F,  BHALUKA"},
    {"id": "1rlw6ZFYCsPavas9I-NW7W5lv4u6v_fgH", "orig_parent": "10sIS-KK8MyWGFs5zY06Wqgqi35vaCxKl", "name": "DPSR,  BHALUKA"},
]

def restore():
    if not os.path.exists(TOKEN_FILE):
        print("token.json not found.")
        return

    creds = Credentials.from_authorized_user_file(
        TOKEN_FILE,
        ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    )
    ds = build('drive', 'v3', credentials=creds)
    print("Connected to Drive API. Restoring extra non-Cash-In-Hand files...")

    for item in RESTORE_MAPPING:
        try:
            ds.files().update(
                fileId=item['id'],
                addParents=item['orig_parent'],
                removeParents=MYM_A_FOLDER_ID,
                fields='id, parents'
            ).execute()
            print(f"[RESTORED] '{item['name']}' returned to original folder {item['orig_parent']}")
        except Exception as e:
            print(f"[ERROR] Failed to restore '{item['name']}': {e}")

    print("\n=== RESTORATION COMPLETE ===")

if __name__ == '__main__':
    restore()
