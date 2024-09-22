from __future__ import print_function

import io
import os.path
import smtplib
import subprocess
from datetime import datetime, timedelta
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import numpy as np
import pandas as pd
from bsc_ops_admin.utils import get_credentials, get_current_semester_year, get_google_services, upload_to_drive
from googleapiclient.http import MediaIoBaseDownload

DOCUMENT_IDS = {
    "potential_termination_reinstatement_eligible": "1nnFIcC3429tWAq9ihA9ZHv86S78zEPKwht8IT3MLios",
    "potential_termination_reinstatement_ineligible": "1GycmYHjtUybWNmv7M8io1umcfIOUjtAqi1MeetLulpI",
    "pending_termination_notice_reinstatement_eligible": "1Dn0sXHYwbj3BDGjWHR6_PIT2BJBooZ5HZyFfnRtijA0",
    "pending_termination_notice_reinstatement_ineligible": "1g-n1s12abiYaBiQGkFbqbyUnrwtdwe3oKcrgvj3ZLuA",
    "conditional_contract": "1buFyhb4m7hXia-6uH5PI7eipaVsKPsv7u-z14OJxnwM",
    "down_hours_spreadsheet": "1WiH7RSZe3pjc87PvF1Qp1gUzqC7aVXxTpEwwnRzXr6Y",
    "15_day_notice_spreadsheet": "16c6VvfbFV3JlnrVL-jw1q_FtHVFql_04Ww__4BsWLjQ",
    "instruction_docs": "1jcCLkLd58psZyZLxnFHAfEpOCuROOslcfz0qMJrAxDI",
}

PDFS_FOLDER_ID = "1At4TzVjKsN2Zv5LKCFpbeu2DiRQ6FXSg"

COURTESY_NOTICE_ACTION = "Courtesy Notice"
POTENTIAL_TERMINATION_ACTION = "Potential Termination Notice"
PENDING_TERMINATION_ACTION = "Pending Termination Notice"

EMAIL_TEMPLATES_SUBJECT_LINES = {
    COURTESY_NOTICE_ACTION: "Down 10+ hours - Courtesy Notice",
    POTENTIAL_TERMINATION_ACTION: "15-Day Notice of Potential Membership Termination",
    PENDING_TERMINATION_ACTION: "[URGENT] 15 Day Notice of Pending Contract and Membership Termination",
}

OPS_SUPERVISOR = "Alex"


SEMESTER_YEAR = get_current_semester_year()
SAMPLE_RANGE_NAME = "sheet1!A:O"
SAFE_MODE = True

COL_LAST_NAME = "C"
COL_FIRST_NAME = "D"
COL_EMAIL = "E"
COL_HOUSE = "F"
COL_EXISTING_CC = "H"
COL_ACTION = "I"


def get_down_hours_df(sheets_service, only_action_null=True):
    down_hours_spreadsheet_id = DOCUMENT_IDS["down_hours_spreadsheet"]
    result = (
        sheets_service.spreadsheets()
        .values()
        .get(spreadsheetId=down_hours_spreadsheet_id, range=SAMPLE_RANGE_NAME)
        .execute()
    )
    values = result.get("values", [])

    df = pd.DataFrame(values)
    df.columns = df.iloc[0]
    df = df.drop(0)

    if only_action_null:
        df = df[df["Action"].isnull()]
        assert len(df) < 5, "Something is fishy, there are more than 5 rows with no action?"

    df["Member's Down Hours"] = df["Member's Down Hours"].replace("", np.nan).astype(float)
    return df


def update_down_hours_spreadsheet_cell(sheets_service, col, idx, value):
    down_hours_spreadsheet_id = DOCUMENT_IDS["down_hours_spreadsheet"]
    return (
        sheets_service.spreadsheets()
        .values()
        .update(
            spreadsheetId=down_hours_spreadsheet_id,
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


def had_prior_conditional_contract(row, full_df):
    member_email = row["Member's Email"].strip()
    assert "@" in member_email, f"Member email {member_email} is not valid"
    to_check = full_df.iloc[0 : row.name]
    # Checking whether the person ever had a potential termination action (i.e. got a conditional contract)
    prior_CC = to_check[
        (to_check["Member's Email"] == member_email) & (to_check["Action"] == POTENTIAL_TERMINATION_ACTION)
    ]
    existing_contract = "Yes" if len(prior_CC) > 0 else "No"
    return existing_contract, member_email


def get_action(row, had_prior_CC):
    down_hours = row["Member's Down Hours"]
    if down_hours >= 15:
        if had_prior_CC == "Yes":
            action = PENDING_TERMINATION_ACTION
        elif had_prior_CC == "No":
            action = POTENTIAL_TERMINATION_ACTION
        else:
            raise ValueError(f"Existing CC should only be 'Yes' or 'No', it is {had_prior_CC}")
    elif down_hours >= 10:
        action = COURTESY_NOTICE_ACTION
    else:
        raise ValueError(f"Down hours should be >= 10 to be on the form, it is {down_hours}")
    return action


def find_email_if_not_found(sheet):
    df = get_down_hours_df(sheet)
    ppl = pd.read_csv("/Users/admin/Downloads/PersonListExport.csv")

    for i in range(0, len(df)):
        if df.iloc[i]["Member's Email"] is None or df.iloc[i]["Member's Email"] == "":
            try:
                first_name = df.iloc[i]["Member's First Name"]
                last_name = df.iloc[i]["Member's Last Name"]
                email = get_email(ppl, first_name, last_name)
                print(email)
                update_down_hours_spreadsheet_cell(sheet, COL_EMAIL, df.iloc[i].name, email)
            except Exception as e:
                print(first_name, last_name, e)
                update_down_hours_spreadsheet_cell(sheet, COL_EMAIL, df.iloc[i].name, "NOT FOUND")


def delete_file(service, file_id):
    try:
        service.files().delete(fileId=file_id).execute()
        print(f"File with ID {file_id} has been deleted.")
    except Exception as e:
        print(f"An error occurred while deleting the file: {e}")


def fill_pdf(services, output_pdf_path, form_data, document_id):
    drive_service, docs_service = services["drive"], services["docs"]

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

    print(f"PDF form filled and saved as {output_pdf_path}")


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


def extract_email_templates(docs_service):
    template_doc = DOCUMENT_IDS["instruction_docs"]

    # Retrieve the document content
    document = docs_service.documents().get(documentId=template_doc).execute()
    content = document.get("body").get("content", [])

    templates = {}
    current_subject = None
    current_template = []
    in_template = False

    for element in content:
        if "paragraph" in element:
            paragraph = element.get("paragraph", {})
            text = "".join([run.get("textRun", {}).get("content", "") for run in paragraph.get("elements", [])])

            if is_roboto_mono(paragraph):
                if not in_template:
                    current_subject = text.split("Subject: ")[1].strip()
                    current_template = []
                    in_template = True
                else:
                    # This is part of the template body
                    current_template.append(text)
            elif in_template:
                templates[current_subject] = "".join(current_template).strip()
                in_template = False

    return templates


def is_roboto_mono(paragraph):
    for run in paragraph.get("elements", []):
        if "textRun" in run and "textStyle" in run["textRun"] and "weightedFontFamily" in run["textRun"]["textStyle"]:
            if run["textRun"]["textStyle"]["weightedFontFamily"]["fontFamily"] == "Roboto Mono":
                return True
    return False


def get_email_template(templates, action):
    subject = EMAIL_TEMPLATES_SUBJECT_LINES[action]
    body = templates.get(subject)
    if body is None:
        raise ValueError(f"Email template for '{subject}' not found")
    return subject, body


def get_reinstatement_eligibility_suffix(member_first_name, member_last_name):
    is_reinstatement_eligible = input(
        f"Is {member_first_name} {member_last_name} eligible for reinstatement? (y/n) (default: y)"
    )
    assert is_reinstatement_eligible in ["y", "n", ""], f"Invalid input {is_reinstatement_eligible}"

    if is_reinstatement_eligible == "n":
        prior_termination_reason = input("Enter the reason for prior termination:")
    else:
        prior_termination_reason = None

    eligibility_suffix = "ineligible" if is_reinstatement_eligible == "n" else "eligible"
    return eligibility_suffix, prior_termination_reason


def update_15_day_notice_spreadsheet(sheets_service, member_data, date_15days):
    spreadsheet_id = DOCUMENT_IDS["15_day_notice_spreadsheet"]

    # First, insert a new row after the header row
    request = {
        "insertDimension": {
            "range": {
                "sheetId": 0,  # Assuming it's the first sheet
                "dimension": "ROWS",
                "startIndex": 2,  # Insert after the header row
                "endIndex": 3,
            },
            "inheritFromBefore": False,
        }
    }
    sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": [request]}).execute()

    # Now, update the newly inserted row with member data
    COLUMNS_BY_DATA = {
        "<LAST NAME>": "B",
        "<FIRST NAME>": "C",
        "<EMAIL>": "D",
        "<HOUSE>": "E",
        "<DATE>": "F",
        "<DATE (+15 days)>": "G",
        "<CURRENT REASON>": "H",
        "<ACTION>": "I",
    }

    # Prepare the update requests
    requests = []
    for data_key, column in COLUMNS_BY_DATA.items():
        if data_key == "<CURRENT REASON>":
            value = "Workshift"
        elif data_key == "<ACTION>":
            action = format_data[data_key]
            if action == POTENTIAL_TERMINATION_ACTION:
                action = "Potential"
            elif action == PENDING_TERMINATION_ACTION:
                action = "Pending"
            else:
                raise ValueError(f"Invalid action {action}")
            value = action
        else:
            value = format_data[data_key]

        if value:
            requests.append(
                {
                    "updateCells": {
                        "range": {
                            "sheetId": 0,
                            "startRowIndex": 2,
                            "endRowIndex": 3,
                            "startColumnIndex": ord(column) - ord("A"),
                            "endColumnIndex": ord(column) - ord("A") + 1,
                        },
                        "rows": [{"values": [{"userEnteredValue": {"stringValue": str(value)}}]}],
                        "fields": "userEnteredValue",
                    }
                }
            )

    # Execute all updates in a single batch request
    if requests:
        sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests}).execute()

    print(f"Updated 15-day notice spreadsheet for {format_data['<FULL NAME>']}")


if __name__ == "__main__":
    creds = get_credentials()
    services = get_google_services(creds)
    full_df = get_down_hours_df(services["sheets"], only_action_null=False)
    df = get_down_hours_df(services["sheets"])
    templates = extract_email_templates(services["docs"])

    print(f"There are {len(df)} rows to process")

    # Example usage
    for i, row in df.iterrows():
        house_code = get_house_code(row)
        member_first_name, member_last_name = get_capitalized_names(row)
        had_prior_CC, member_email = had_prior_conditional_contract(row, full_df)
        action = get_action(row, had_prior_CC)
        workshift_manager_email = row["Email Address"].strip()

        print(f"Action for {member_first_name} {member_last_name} is {action}")
        date_today = datetime.now().strftime("%m/%d/%Y")
        date_7days = (datetime.now() + timedelta(days=7)).strftime("%m/%d/%Y")
        date_15days = (datetime.now() + timedelta(days=15)).strftime("%m/%d/%Y")
        format_data = {
            "<FIRST NAME>": member_first_name,
            "<LAST NAME>": member_last_name,
            "<FULL NAME>": f"{member_first_name} {member_last_name}",
            "<HOUSE>": house_code,
            "<DATE>": date_today,
            "<DATE (+1 week)>": date_7days,
            "<DATE (+15 days)>": date_15days,
            "<SEMESTER, YEAR>": SEMESTER_YEAR,
            "<OPS_SUPERVISOR>": OPS_SUPERVISOR,
            "<EMAIL>": member_email,
            "<ACTION>": action,
        }

        if action == POTENTIAL_TERMINATION_ACTION:
            # Get from input whether member is eligible for reinstatement, default to eligible
            eligibility_suffix, prior_termination_reason = get_reinstatement_eligibility_suffix(
                member_first_name, member_last_name
            )
            if prior_termination_reason is not None:
                format_data["<PRIOR TERMINATION REASON>"] = prior_termination_reason

            document_id = DOCUMENT_IDS["conditional_contract"]
            cc_pdf = f"cc_{member_first_name}_{member_last_name}.pdf"
            fill_pdf(services, cc_pdf, format_data, document_id)

            document_id = DOCUMENT_IDS[f"potential_termination_reinstatement_{eligibility_suffix}"]
            potential_termination_pdf = f"potential_termination_{member_first_name}_{member_last_name}.pdf"
            fill_pdf(services, potential_termination_pdf, format_data, document_id)

            pdf_attachments = [cc_pdf, potential_termination_pdf]

            subject, body = get_email_template(templates, action)

        elif action == PENDING_TERMINATION_ACTION:
            eligibility_suffix, prior_termination_reason = get_reinstatement_eligibility_suffix(
                member_first_name, member_last_name
            )
            if prior_termination_reason is not None:
                format_data["<PRIOR TERMINATION REASON>"] = prior_termination_reason

            document_id = DOCUMENT_IDS[f"pending_termination_notice_reinstatement_{eligibility_suffix}"]
            pending_termination_pdf = f"pending_termination_{member_first_name}_{member_last_name}.pdf"
            fill_pdf(services, pending_termination_pdf, format_data, document_id)

            pdf_attachments = [pending_termination_pdf]

            subject, body = get_email_template(templates, action)

        elif action == COURTESY_NOTICE_ACTION:
            pdf_attachments = []

            subject, body = get_email_template(templates, action)

        # Format the body. Wrap all < > with {} in body and then use formatting
        body = body.replace("<", "{<").replace(">", ">}").format_map(format_data)

        if SAFE_MODE:
            if action == POTENTIAL_TERMINATION_ACTION:
                open_pdf_in_preview(cc_pdf)
                open_pdf_in_preview(potential_termination_pdf)
            elif action == PENDING_TERMINATION_ACTION:
                open_pdf_in_preview(pending_termination_pdf)
            input(
                f"About to send email to {member_email} and {workshift_manager_email}.\n\nSubject: {subject}\n\nBody:\n{body}\n\nAttachments: {', '.join(os.path.basename(path) for path in pdf_attachments)}\n\nPress Enter to continue"
            )

        assert "{" not in body and "}" not in body, "Body has { or } in it, probably formatting failed"
        assert "<" not in body and ">" not in body, "Body has < or > in it, which are not formatted properly"

        ####################################################
        # Start doing the actual work that modifies things #
        ####################################################

        # Update 15 day notice spreadsheet
        if action == POTENTIAL_TERMINATION_ACTION or action == PENDING_TERMINATION_ACTION:
            print(f"Updating 15 day notice spreadsheet for {member_first_name} {member_last_name}")
            update_15_day_notice_spreadsheet(services["sheets"], format_data, date_15days)

        for pdf_path in pdf_attachments:
            print(f"Uploading {pdf_path} to Google Drive")
            drive_file_id = upload_to_drive(services["drive"], pdf_path, PDFS_FOLDER_ID)
            print(f"Uploaded {pdf_path} to Google Drive with file ID: {drive_file_id}")

        # Actually send the email
        print(f"Sending email to {member_email} and {workshift_manager_email}")
        send_email(member_email, [workshift_manager_email], subject, body, pdf_attachments)
        print("Email sent.")

        # Update down hours spreadsheet
        spreadsheet_items_to_update = {
            COL_HOUSE: house_code,
            COL_LAST_NAME: member_last_name,
            COL_FIRST_NAME: member_first_name,
            COL_EXISTING_CC: had_prior_CC,
            COL_ACTION: action,
        }
        print(f"About to update spreadsheet values: {spreadsheet_items_to_update}")

        for col, value in spreadsheet_items_to_update.items():
            print(f"Updating {member_first_name} {member_last_name} col {col} to {value}")
            update_down_hours_spreadsheet_cell(services["sheets"], col, row.name, value)

        print(f"Finished everything for {member_first_name} {member_last_name}")
