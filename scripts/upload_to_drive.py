"""
Upload Marketing_Dashboard.html to Google Drive.

First-time setup:
1. Go to https://console.cloud.google.com/
2. Create a project, enable Google Drive API
3. Create OAuth 2.0 credentials (Desktop app)
4. Download as credentials.json and place next to this file
5. Run this script once manually — it will open a browser to authorize
6. After auth, token.json is saved and future runs are fully automatic
"""

import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent.parent / ".env")

DASHBOARD = Path(__file__).parent.parent / "Marketing_Dashboard.html"
CREDS_FILE = Path(__file__).parent / "credentials.json"
TOKEN_FILE = Path(__file__).parent / "token.json"
GDRIVE_FILE_ID = os.getenv("GDRIVE_FILE_ID", "")

SCOPES = ["https://www.googleapis.com/auth/drive.file"]


def get_service():
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDS_FILE.exists():
                raise FileNotFoundError(
                    f"credentials.json not found at {CREDS_FILE}\n"
                    "See script header for setup instructions."
                )
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDS_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_FILE.write_text(creds.to_json())

    return build("drive", "v3", credentials=creds)


def upload():
    from googleapiclient.http import MediaFileUpload

    service = get_service()
    media = MediaFileUpload(
        str(DASHBOARD),
        mimetype="text/html",
        resumable=False,
    )

    if GDRIVE_FILE_ID:
        # Update existing file — rename it to include today's date
        from datetime import date
        new_name = f"Marketing_Dashboard_{date.today().strftime('%Y-%m-%d')}.html"
        file = service.files().update(
            fileId=GDRIVE_FILE_ID,
            body={"name": new_name},
            media_body=media,
        ).execute()
        print(f"Updated file in Drive: {file.get('name')} (id: {file.get('id')})")
    else:
        # Create new file
        meta = {"name": "Marketing_Dashboard.html"}
        file = service.files().create(
            body=meta,
            media_body=media,
            fields="id,name,webViewLink",
        ).execute()
        file_id = file.get("id")
        print(f"Created file in Drive: {file.get('name')}")
        print(f"  File ID: {file_id}")
        print(f"  View:    {file.get('webViewLink')}")
        print(f"\nSave this File ID to .env as GDRIVE_FILE_ID={file_id}")


if __name__ == "__main__":
    upload()
