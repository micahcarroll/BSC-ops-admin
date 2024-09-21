from __future__ import print_function

import io
import os.path
import pickle
import smtplib
import subprocess
from datetime import datetime, timedelta
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

import numpy as np
import pandas as pd
from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

SPREADSHEET_ID = "1WiH7RSZe3pjc87PvF1Qp1gUzqC7aVXxTpEwwnRzXr6Y"
POTENTIAL_TERMINATION_NOTICE_DOCUMENT_ID = "1nnFIcC3429tWAq9ihA9ZHv86S78zEPKwht8IT3MLios"
CONDITIONAL_CONTRACT_DOCUMENT_ID = "1buFyhb4m7hXia-6uH5PI7eipaVsKPsv7u-z14OJxnwM"
OPS_SUPERVISOR = "Alex"
SEMESTER_YEAR = "Fall 2024"  # TODO: make this dynamic
SAMPLE_RANGE_NAME = "sheet1!A:O"
SAFE_MODE = True

ENV_FOLDER = Path(__file__).parent / ".env"
TEMPLATE_FOLDER = Path(__file__).parent / "templates"

COL_LAST_NAME = "C"
COL_FIRST_NAME = "D"
COL_EMAIL = "E"
COL_HOUSE = "F"
COL_EXISTING_CC = "H"
COL_ACTION = "I"

POTENTIAL_TERMINATION_ACTION = "Potential Termination Notice"
PENDING_TERMINATION_ACTION = "Pending Termination Notice"
COURTESY_NOTICE_ACTION = "Courtesy Notice"


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


def get_sheet(creds):
    service = build("sheets", "v4", credentials=creds)
    return service.spreadsheets()


def get_df(sheet, only_action_null=True):
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute()
    values = result.get("values", [])

    df = pd.DataFrame(values)
    df.columns = df.iloc[0]
    df = df.drop(0)

    if only_action_null:
        df = df[df["Action"].isnull()]
        assert len(df) < 5, "Something is fishy, there are more than 5 rows with no action?"

    df["Member's Down Hours"] = df["Member's Down Hours"].replace("", np.nan).astype(float)
    return df


def update_cell(sheet, col, idx, value):
    return (
        sheet.values()
        .update(
            spreadsheetId=SPREADSHEET_ID,
            range=f"sheet1!{col}{idx+1}",
            valueInputOption="USER_ENTERED",
            body={"values": [[value]], "majorDimension": "COLUMNS"},
        )
        .execute()
    )


def get_house_code(row):
    house_code = row["Email Address"][0:3].upper().strip()
    return house_code


def get_email(ppl, first_name, last_name):
    return ppl.loc[(ppl["Last Name"] == last_name) & (ppl["First Name"] == first_name)]["Permanent Email"].values[0]


def get_capitalized_names(row):
    first_name = row["Member's First Name"]
    last_name = row["Member's Last Name"]
    return first_name.title(), last_name.title()


def has_existing_conditional_contract(row, full_df):
    member_email = row["Member's Email"].strip()
    assert "@" in member_email, f"Member email {member_email} is not valid"
    to_check = full_df.iloc[0 : row.name]
    # TODO: is this actually the right way to do this?
    prior_large_down_hours = to_check[
        (to_check["Member's Email"] == member_email) & (to_check["Member's Down Hours"] >= 15)
    ]
    existing_contract = "Yes" if len(prior_large_down_hours) >= 2 else "No"
    return existing_contract, member_email


def get_action(row, existing_cc):
    down_hours = row["Member's Down Hours"]
    if float(down_hours) >= 15:
        if existing_cc == "Yes":
            action = PENDING_TERMINATION_ACTION
        elif existing_cc == "No":
            action = POTENTIAL_TERMINATION_ACTION
        else:
            raise ValueError(f"Existing CC should only be 'Yes' or 'No', it is {existing_cc}")
    else:
        action = COURTESY_NOTICE_ACTION
    return action


def find_email_if_not_found(sheet):
    df = get_df(sheet)
    ppl = pd.read_csv("/Users/admin/Downloads/PersonListExport.csv")

    for i in range(0, len(df)):
        if df.iloc[i]["Member's Email"] is None or df.iloc[i]["Member's Email"] == "":
            try:
                first_name = df.iloc[i]["Member's First Name"]
                last_name = df.iloc[i]["Member's Last Name"]
                email = get_email(ppl, first_name, last_name)
                print(email)
                update_cell(sheet, COL_EMAIL, df.iloc[i].name, email)
            except Exception as e:
                print(first_name, last_name, e)
                update_cell(sheet, COL_EMAIL, df.iloc[i].name, "NOT FOUND")


def delete_file(service, file_id):
    try:
        service.files().delete(fileId=file_id).execute()
        print(f"File with ID {file_id} has been deleted.")
    except Exception as e:
        print(f"An error occurred while deleting the file: {e}")


def fill_pdf(output_pdf_path, form_data, document_id):
    creds = get_credentials()

    # Build the Google Docs and Drive services
    docs_service = build("docs", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)

    # Create a copy of the document
    body = {"name": "tmp"}
    drive_response = drive_service.files().copy(fileId=document_id, body=body).execute()
    copy_document_id = drive_response.get("id")

    print(f"Created a copy of the document with ID: {copy_document_id}")

    # Retrieve the documents contents from the Docs service.
    docs_service.documents().get(documentId=copy_document_id).execute()

    # Update the document content
    requests = []
    for key, value in form_data.items():
        requests.append({"replaceAllText": {"containsText": {"text": key, "matchCase": "true"}, "replaceText": value}})

    # Execute the update
    docs_service.documents().batchUpdate(documentId=copy_document_id, body={"requests": requests}).execute()

    # Export the document as PDF
    request = drive_service.files().export_media(fileId=copy_document_id, mimeType="application/pdf")
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print(f"Download {int(status.progress() * 100)}%.")

    delete_file(drive_service, copy_document_id)

    # Save the PDF
    fh.seek(0)
    with open(output_pdf_path, "wb") as f:
        f.write(fh.getvalue())


def send_email(recipient_email, cc_emails, subject, body, attachment_paths):
    # Email configuration
    sender_email = "opsadmin@bsc.coop"
    # Get from environment variable
    sender_password = os.getenv("EMAIL_PASSWORD")

    # Create the email message
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = recipient_email
    message["Subject"] = subject
    message["Cc"] = ", ".join(cc_emails)

    # Add body to email
    message.attach(MIMEText(body, "plain"))

    # Attach multiple files
    for attachment_path in attachment_paths:
        with open(attachment_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename= {os.path.basename(attachment_path)}")
        message.attach(part)

    # Connect to the SMTP server and send the email
    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)  # type: ignore
            server.send_message(message)
        print(f"Email sent successfully to {recipient_email} with CC to {', '.join(cc_emails)}")
        print(f"Attachments: {', '.join(os.path.basename(path) for path in attachment_paths)}")
    except Exception as e:
        print(f"Error sending email: {e}")


def open_pdf_in_preview(pdf_path):
    try:
        # Use the 'open' command to open the PDF in Preview
        subprocess.run(["open", "-a", "Preview", pdf_path], check=True)
        print(f"Opened {pdf_path} in Preview")
    except subprocess.CalledProcessError as e:
        print(f"Error opening PDF: {e}")
    except FileNotFoundError:
        print(f"File not found: {pdf_path}")


if __name__ == "__main__":
    creds = get_credentials()
    sheet = get_sheet(creds)
    full_df = get_df(sheet, only_action_null=False)
    df = get_df(sheet)

    # If save mode, ask user to confirm if the amount seems right
    if SAFE_MODE:
        input(f"There are {len(df)} unprocessed entries in the spreadsheet. Press Enter to continue")

    # Example usage
    for i, row in df.iterrows():
        house_code = get_house_code(row)
        member_first_name, member_last_name = get_capitalized_names(row)
        existing_contract, member_email = has_existing_conditional_contract(row, full_df)
        action = get_action(row, existing_contract)
        workshift_manager_email = row["Email Address"].strip()
        print(f"Action for {member_first_name} {member_last_name} is {action}")
        date_today = datetime.now().strftime("%m/%d/%Y")
        date_7days = (datetime.now() + timedelta(days=7)).strftime("%m/%d/%Y")
        date_15days = (datetime.now() + timedelta(days=15)).strftime("%m/%d/%Y")

        if action == POTENTIAL_TERMINATION_ACTION:

            form_data = {
                "<FIRST NAME>": member_first_name,
                "<FULL NAME>": f"{member_first_name} {member_last_name}",
                "<HOUSE>": house_code,
                "<DATE>": date_today,
                "<DATE (+1 week)>": date_7days,
                "<DATE (+15 days)>": date_15days,
                "<SEMESTER, YEAR>": SEMESTER_YEAR,
            }

            cc_output_pdf = f"cc_{member_first_name}_{member_last_name}.pdf"
            fill_pdf(cc_output_pdf, form_data, CONDITIONAL_CONTRACT_DOCUMENT_ID)
            print(f"PDF form filled and saved as {cc_output_pdf}")

            potential_termination_output_pdf = f"potential_termination_{member_first_name}_{member_last_name}.pdf"
            fill_pdf(potential_termination_output_pdf, form_data, POTENTIAL_TERMINATION_NOTICE_DOCUMENT_ID)
            print(f"PDF form filled and saved as {potential_termination_output_pdf}")

            pdf_attachments = [cc_output_pdf, potential_termination_output_pdf]

            subject = "15-Day Notice of Potential Membership Termination"
            # Load the email template from the file
            with open(TEMPLATE_FOLDER / "potential_termination_email_template.txt", "r") as file:
                body = file.read().format_map(
                    {
                        "first_name": member_first_name,
                        "date_15days": date_15days,
                        "date_7days": date_7days,
                        "OPS_SUPERVISOR": OPS_SUPERVISOR,
                    }
                )
        elif action == PENDING_TERMINATION_ACTION:
            subject = "[URGENT] 15 Day Notice of Pending Contract and Membership Termination"
            with open(TEMPLATE_FOLDER / "pending_termination_email_template.txt", "r") as file:
                body = file.read().format_map(locals())
        elif action == COURTESY_NOTICE_ACTION:
            subject = "Down 10+ hours - Courtesy Notice"
            with open(TEMPLATE_FOLDER / "courtesy_notice_email_template.txt", "r") as file:
                body = file.read().format_map(locals())

        if SAFE_MODE:
            open_pdf_in_preview(cc_output_pdf)
            open_pdf_in_preview(potential_termination_output_pdf)
            input(
                f"About to send email to {member_email} and {workshift_manager_email}.\n\nSubject: {subject}\n\nBody:\n{body}\n\nAttachments: {', '.join(os.path.basename(path) for path in pdf_attachments)}\n\nPress Enter to continue"
            )

        send_email(member_email, [workshift_manager_email], subject, body, pdf_attachments)

        spreadsheet_items_to_update = {
            COL_HOUSE: house_code,
            COL_LAST_NAME: member_last_name,
            COL_FIRST_NAME: member_first_name,
            COL_EXISTING_CC: existing_contract,
            COL_ACTION: action,
        }
        print(f"Updated spreadsheet values: {spreadsheet_items_to_update}")

        for col, value in spreadsheet_items_to_update.items():
            print(f"Updating {member_first_name} {member_last_name} col {col} to {value}")
            update_cell(sheet, col, row.name, value)
