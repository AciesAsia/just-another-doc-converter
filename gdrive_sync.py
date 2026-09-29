"""
gdrive_sync.py
--------------
Downloads all PDF files from a Google Drive folder to a local directory,
then deletes them from Drive on successful download.

Setup (one-time):
  1. Go to https://console.cloud.google.com/
  2. Create a project, enable "Google Drive API"
  3. Create OAuth 2.0 credentials (Desktop app), download as credentials.json
  4. Place credentials.json in the same folder as this script
  5. Run the script — a browser window opens for Google sign-in (once only)

Usage:
  python gdrive_sync.py
"""

import os
import sys
import json
from pathlib import Path

# ── Configuration ─────────────────────────────────────────────────────────────

DRIVE_FOLDER_ID = "19RVhksgAetnFtdjQdT7NEgTn3OzI8yfk"

LOCAL_DEST = Path(r"C:\Users\User\Documents\PROJECTS\LRM_DFS\02_PROJECT_FILES\Source_Docs")

SCOPES = ["https://www.googleapis.com/auth/drive"]

SCRIPT_DIR = Path(__file__).parent
CREDENTIALS_FILE = SCRIPT_DIR / "credentials.json"
TOKEN_FILE = SCRIPT_DIR / "token.json"

# ── Auth ──────────────────────────────────────────────────────────────────────

def get_drive_service():
    try:
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from google.auth.transport.requests import Request
        from googleapiclient.discovery import build
    except ImportError:
        print("[ERROR] Missing dependencies. Run:")
        print("  pip install google-api-python-client google-auth-oauthlib google-auth-httplib2")
        sys.exit(1)

    if not CREDENTIALS_FILE.exists():
        print(f"[ERROR] credentials.json not found at: {CREDENTIALS_FILE}")
        print("Download it from Google Cloud Console > APIs & Services > Credentials")
        sys.exit(1)

    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_FILE.write_text(creds.to_json())

    return build("drive", "v3", credentials=creds)

# ── Core logic ────────────────────────────────────────────────────────────────

def list_pdfs(service):
    query = f"'{DRIVE_FOLDER_ID}' in parents and mimeType='application/pdf' and trashed=false"
    results = service.files().list(
        q=query,
        fields="files(id, name, size)",
        pageSize=1000
    ).execute()
    return results.get("files", [])


def download_file(service, file_id, file_name, dest_dir):
    from googleapiclient.http import MediaIoBaseDownload
    import io

    dest_path = dest_dir / file_name
    request = service.files().get_media(fileId=file_id)
    with open(dest_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
    return dest_path


def delete_file(service, file_id):
    service.files().delete(fileId=file_id).execute()


def run():
    print("=" * 55)
    print("  Google Drive PDF Sync")
    print("=" * 55)

    # Ensure local destination exists
    LOCAL_DEST.mkdir(parents=True, exist_ok=True)

    service = get_drive_service()

    files = list_pdfs(service)
    if not files:
        print("No PDF files found in the Drive folder.")
        print("=" * 55)
        return

    print(f"Found {len(files)} PDF(s) to download.\n")

    downloaded = 0
    failed = 0

    for f in files:
        name = f["name"]
        fid = f["id"]
        size_kb = int(f.get("size", 0)) // 1024
        print(f"  Downloading: {name} ({size_kb} KB) ... ", end="", flush=True)
        try:
            dest = download_file(service, fid, name, LOCAL_DEST)
            # Verify file landed on disk
            if dest.exists() and dest.stat().st_size > 0:
                delete_file(service, fid)
                print("OK (deleted from Drive)")
                downloaded += 1
            else:
                print("FAILED (empty file, kept on Drive)")
                failed += 1
        except Exception as e:
            print(f"FAILED ({e})")
            failed += 1

    print()
    print(f"  Downloaded : {downloaded}")
    print(f"  Failed     : {failed}")
    print(f"  Saved to   : {LOCAL_DEST}")
    print("=" * 55)


if __name__ == "__main__":
    run()
