import os
import pickle
from datetime import datetime
from pathlib import Path
from urllib.error import HTTPError

from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

ENV_FOLDER = Path(__file__).parent / ".env"
TEMPLATE_FOLDER = Path(__file__).parent / "templates"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# Load environment variables from .env file
load_dotenv(ENV_FOLDER / ".env")


def get_credentials():
    creds = None
    if os.path.exists(ENV_FOLDER / "token.pickle"):
        with open(ENV_FOLDER / "token.pickle", "rb") as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(ENV_FOLDER / "credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open(ENV_FOLDER / "token.pickle", "wb") as token:
            pickle.dump(creds, token)
    return creds


def get_google_services(creds):
    sheet_service = build("sheets", "v4", credentials=creds)
    docs_service = build("docs", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)
    return {"sheets": sheet_service, "docs": docs_service, "drive": drive_service}


def upload_to_drive(drive_service, file_path, folder_id):
    file_metadata = {"name": os.path.basename(file_path), "parents": [folder_id]}
    media = MediaFileUpload(file_path, resumable=True)

    try:
        # First, create the file metadata
        file = drive_service.files().create(body=file_metadata, fields="id", supportsAllDrives=True).execute()

        # Then, upload the file content
        drive_service.files().update(fileId=file.get("id"), media_body=media, supportsAllDrives=True).execute()

        print(f'File uploaded successfully. File ID: {file.get("id")}')
        return file.get("id")
    except HTTPError as error:
        print(f"An error occurred while uploading the file: {error}")
        return None


def get_current_semester_year():
    current_date = datetime.now()

    # Define cutoff dates for semesters
    spring_start = datetime(current_date.year, 1, 1)
    summer_start = datetime(current_date.year, 5, 15)
    fall_start = datetime(current_date.year, 8, 15)

    if spring_start <= current_date < summer_start:
        semester = "Spring"
    elif summer_start <= current_date < fall_start:
        semester = "Summer"
    else:
        semester = "Fall"

    # Adjust year for Fall semester
    if semester == "Fall" and current_date.month >= 8:
        year = current_date.year + 1
    else:
        year = current_date.year

    return f"{semester} {year}"
