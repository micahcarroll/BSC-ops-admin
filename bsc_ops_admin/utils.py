import os
import pickle
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

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
