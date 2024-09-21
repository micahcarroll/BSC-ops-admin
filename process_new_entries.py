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

import pandas as pd
from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1WiH7RSZe3pjc87PvF1Qp1gUzqC7aVXxTpEwwnRzXr6Y"
POTENTIAL_TERMINATION_NOTICE_DOCUMENT_ID = "1nnFIcC3429tWAq9ihA9ZHv86S78zEPKwht8IT3MLios"
CONDITIONAL_CONTRACT_DOCUMENT_ID = "1buFyhb4m7hXia-6uH5PI7eipaVsKPsv7u-z14OJxnwM"
OPS_SUPERVISOR = "Alex"
SEMESTER_YEAR = "Fall 2024"  # TODO: make this dynamic
SAMPLE_RANGE_NAME = "sheet1!A:O"
SAFE_MODE = True

ENV_FOLDER = Path(__file__).parent.parent / ".env"
TEMPLATE_FOLDER = Path(__file__).parent.parent / "templates"

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
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute()
    values = result.get("values", [])

    df = pd.DataFrame(values)
    df.columns = df.iloc[0]
    df = df.drop(0)

    if only_action_null:
        df = df[df["Action"].isnull()]
        assert len(df) < 5, "Something is fishy, there are more than 5 rows with no action?"

    return df


def update_cell(sheet, col, idx, value):
    return (
        sheet.values()
        .update(
            spreadsheetId=SAMPLE_SPREADSHEET_ID,
            range=f"sheet1!{col}{idx+1}",
            valueInputOption="USER_ENTERED",
            body={"values": [[value]], "majorDimension": "COLUMNS"},
        )
        .execute()
    )


def fill_house(sheet):
    df = get_df(sheet)
    for i in range(0, len(df)):
        entry = df.iloc[i]
        if entry["House"].strip() == "":
            first_name = entry["Member's First Name"]
            last_name = entry["Member's Last Name"]
            house_code = entry["Email Address"][0:3].upper()
            print(f"Updating house for {first_name} {last_name} to {house_code}")
            update_cell(sheet, "F", entry.name, house_code)


def get_email(ppl, first_name, last_name):
    return ppl.loc[(ppl["Last Name"] == last_name) & (ppl["First Name"] == first_name)]["Permanent Email"].values[0]


def capitalize_names(sheet):
    df = get_df(sheet)
    for i in range(0, len(df)):
        entry = df.iloc[i]
        first_name = entry["Member's First Name"]
        last_name = entry["Member's Last Name"]
        if last_name != last_name.title():
            print(f"Updating last name for {first_name} {last_name} to {last_name.title()}")
            update_cell(sheet, "C", entry.name, last_name.title())
        if first_name != first_name.title():
            print(f"Updating last name for {first_name} {last_name} to {last_name.title()}")
            update_cell(sheet, "D", entry.name, first_name.title())


def fill_existing_contract(sheet):
    df = get_df(sheet)
    for i in range(0, len(df)):
        entry = df.iloc[i]
        member_email = entry["Member's Email"]
        member_first_name = entry["Member's First Name"]
        member_last_name = entry["Member's Last Name"]

        if member_email is not None:
            print(member_email)
            if "@" in member_email:
                full_df = get_df(sheet, only_action_null=False)
                to_check = full_df.iloc[0 : entry.name]
                # TODO: is this actually the right way to do this?
                to_check = to_check.query("`Member's Email` == @member_email")
                to_check["Member's Down Hours"] = to_check["Member's Down Hours"].astype(float)
                to_check = to_check.query("`Member's Down Hours` >= 15")
                if len(to_check) >= 2:
                    assert False, "This part of the code has not been tested"
                    existing_contract = "Yes"
                else:
                    existing_contract = "No"
            # else:
            #     assert False, "This part of the code has not been tested"
            #     to_check = df.iloc[0 : entry.name]
            #     to_check = to_check.loc[
            #         (to_check["Member's Down Hours"] >= 15)
            #         & (to_check["Member's First Name"] == member_first_name)
            #         & (to_check["Member's Last Name"] == member_last_name)
            #     ]
            #     if len(to_check) >= 2:
            #         existing_contract = "Yes"
            #     else:
            #         existing_contract = "No"

            print(f"Updating existing CC cell for {member_first_name} {member_last_name} to {existing_contract}")
            update_cell(sheet, "H", entry.name, existing_contract)


def fill_action(sheet):
    """Updates the spreadsheet to say that the action has been taken"""
    df = get_df(sheet)
    for i in range(0, len(df)):
        entry = df.iloc[i]
        action = entry["Action"]
        down_hours = entry["Member's Down Hours"]
        existing_cc = entry["Existing CC"]
        if action is None or action == "":
            if float(down_hours) >= 15:
                assert False, "This part of the code has not been tested"
                if existing_cc.strip() == "Yes":
                    action = "Pending Termination Notice"
                elif existing_cc.strip() == "No":
                    action = "Potential Termination Notice"
                else:
                    action = ""
            else:
                action = "Courtesy Notice"
            update_cell(sheet, "I", entry.name, action)


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
                update_cell(sheet, "E", df.iloc[i].name, email)
            except Exception as e:
                print(first_name, last_name, e)
                update_cell(sheet, "E", df.iloc[i].name, "NOT FOUND")


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
    fill_house(sheet)
    capitalize_names(sheet)
    fill_existing_contract(sheet)

    df = get_df(sheet)
    # If save mode, ask user to confirm if the amount seems right
    if SAFE_MODE:
        input(f"There are {len(df)} unprocessed entries in the spreadsheet. Press Enter to continue")

    # Example usage
    for i in range(0, len(df)):
        entry = df.iloc[i]
        first_name = entry["Member's First Name"].strip()
        last_name = entry["Member's Last Name"].strip()
        house = entry["House"].strip()
        member_email = entry["Member's Email"].strip()
        workshift_manager_email = entry["Email Address"].strip()
        date_today = datetime.now().strftime("%m/%d/%Y")
        date_7days = (datetime.now() + timedelta(days=7)).strftime("%m/%d/%Y")
        date_15days = (datetime.now() + timedelta(days=15)).strftime("%m/%d/%Y")
        form_data = {
            "<FIRST NAME>": first_name,
            "<FULL NAME>": f"{first_name} {last_name}",
            "<HOUSE>": house,
            "<DATE>": date_today,
            "<DATE (+1 week)>": date_7days,
            "<DATE (+15 days)>": date_15days,
            "<SEMESTER, YEAR>": SEMESTER_YEAR,
        }

        cc_output_pdf = f"cc_{first_name}_{last_name}.pdf"
        fill_pdf(cc_output_pdf, form_data, CONDITIONAL_CONTRACT_DOCUMENT_ID)
        print(f"PDF form filled and saved as {cc_output_pdf}")

        potential_termination_output_pdf = f"potential_termination_{first_name}_{last_name}.pdf"
        fill_pdf(potential_termination_output_pdf, form_data, POTENTIAL_TERMINATION_NOTICE_DOCUMENT_ID)
        print(f"PDF form filled and saved as {potential_termination_output_pdf}")

        pdf_attachments = [cc_output_pdf, potential_termination_output_pdf]

        subject = "15-Day Notice of Potential Membership Termination"
        # Load the email template from the file
        with open(TEMPLATE_FOLDER / "cc_email_template.txt", "r") as file:
            body = file.read().format_map(locals())

        if SAFE_MODE:
            open_pdf_in_preview(cc_output_pdf)
            open_pdf_in_preview(potential_termination_output_pdf)
            input(
                f"About to send email to {member_email} and {workshift_manager_email}.\n\nSubject: {subject}\n\nBody:\n{body}\n\nAttachments: {', '.join(os.path.basename(path) for path in pdf_attachments)}\n\nPress Enter to continue"
            )
        send_email(member_email, [workshift_manager_email], subject, body, pdf_attachments)

        # This should only be done once the email is sent
        fill_action(sheet)
