"""
gdrive_sync.py
--------------
Converts PDF files in a Google Drive folder to .docx (via Google Docs),
downloads them to a local directory, then deletes all temp files from Drive.

Setup (one-time):
  1. Go to https://console.cloud.google.com/
  2. Create a project, enable "Google Drive API"
  3. Create OAuth 2.0 credentials (Desktop app), download as credentials.json
  4. Place credentials.json in the same folder as this script
  5. Run the script — a browser window opens for Google sign-in (once only)

Usage:
  python gdrive_sync.py
"""

import sys
from pathlib import Path

# ── Configuration ─────────────────────────────────────────────────────────────

SCOPES = ["https://www.googleapis.com/auth/drive"]

if getattr(sys, "frozen", False):
    SCRIPT_DIR = Path(sys.executable).parent
else:
    SCRIPT_DIR = Path(__file__).parent
CREDENTIALS_FILE = SCRIPT_DIR / "credentials.json"
TOKEN_FILE       = SCRIPT_DIR / "token.json"

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
GDOC_MIME = "application/vnd.google-apps.document"

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

# ── Helpers ───────────────────────────────────────────────────────────────────

def extract_folder_id(value):
    """Accept a raw folder ID or a full Drive URL and return just the ID."""
    value = value.strip()
    if "drive.google.com" in value:
        # e.g. https://drive.google.com/drive/folders/<ID>?...
        for part in value.split("/"):
            part = part.split("?")[0]
            if len(part) > 20 and part not in ("drive", "folders", "u", "0"):
                return part
    return value

# ── Core logic ────────────────────────────────────────────────────────────────

def list_pdfs(service, folder_id):
    query = f"'{folder_id}' in parents and mimeType='application/pdf' and trashed=false"
    results = service.files().list(
        q=query,
        fields="files(id, name, size)",
        pageSize=1000
    ).execute()
    return results.get("files", [])


def convert_to_gdoc(service, pdf_id, name, folder_id):
    stem = Path(name).stem
    body = {
        "name": stem,
        "mimeType": GDOC_MIME,
        "parents": [folder_id],
    }
    gdoc = service.files().copy(
        fileId=pdf_id,
        body=body,
        supportsAllDrives=True,
    ).execute()
    return gdoc["id"]


def export_as_docx(service, gdoc_id, stem, dest_dir):
    from googleapiclient.http import MediaIoBaseDownload

    dest_path = dest_dir / f"{stem}.docx"
    request = service.files().export_media(fileId=gdoc_id, mimeType=DOCX_MIME)
    with open(dest_path, "wb") as fh:
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
    return dest_path


def delete_file(service, file_id):
    service.files().delete(fileId=file_id).execute()


def run(drive_folder_id, local_dest):
    local_dest = Path(local_dest)
    folder_id  = extract_folder_id(drive_folder_id)

    print("=" * 55)
    print("  Google Drive PDF → Word Sync")
    print("=" * 55)

    local_dest.mkdir(parents=True, exist_ok=True)

    service = get_drive_service()

    files = list_pdfs(service, folder_id)
    if not files:
        print("No PDF files found in the Drive folder.")
        print("=" * 55)
        return

    print(f"Found {len(files)} PDF(s) to convert.\n")

    converted = 0
    failed    = 0

    for f in files:
        name    = f["name"]
        pdf_id  = f["id"]
        stem    = Path(name).stem
        size_kb = int(f.get("size", 0)) // 1024
        print(f"  Converting: {name} ({size_kb} KB) ... ", end="", flush=True)

        gdoc_id = None
        try:
            gdoc_id = convert_to_gdoc(service, pdf_id, name, folder_id)
            dest    = export_as_docx(service, gdoc_id, stem, local_dest)

            if dest.exists() and dest.stat().st_size > 0:
                delete_file(service, gdoc_id)
                delete_file(service, pdf_id)
                print("OK → .docx saved, Drive cleaned")
                converted += 1
            else:
                print("FAILED (empty output)")
                if gdoc_id:
                    delete_file(service, gdoc_id)
                failed += 1

        except Exception as e:
            print(f"FAILED ({e})")
            if gdoc_id:
                try:
                    delete_file(service, gdoc_id)
                except Exception:
                    pass
            failed += 1

    print()
    print(f"  Converted  : {converted}")
    print(f"  Failed     : {failed}")
    print(f"  Saved to   : {local_dest}")
    print("=" * 55)


if __name__ == "__main__":
    # Defaults when run directly from command line
    _FOLDER_ID  = "19RVhksgAetnFtdjQdT7NEgTn3OzI8yfk"
    _LOCAL_DEST = r"C:\Users\User\Documents\PROJECTS\LRM_DFS\02_PROJECT_FILES\Source_Docs"
    run(_FOLDER_ID, _LOCAL_DEST)
