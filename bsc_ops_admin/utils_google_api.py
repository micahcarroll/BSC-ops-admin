import io

from googleapiclient.http import MediaIoBaseDownload

from bsc_ops_admin.utils import retry_google_api


# HACK: probably consistently failing to delete the files as of right now because of changes to the google drive api
def delete_drive_file(drive_service, file_id, verbose=True):
    if verbose:
        print(f"Deleting file with ID: {file_id}")
    try:
        drive_service.files().delete(fileId=file_id, supportsAllDrives=True).execute()
        if verbose:
            print(f"File with ID {file_id} has been deleted.")
    except Exception as e:
        if "File not found" in str(e) or "notFound" in str(e):
            if verbose:
                print(f"File with ID {file_id} was already deleted or not found.")
        else:
            # Only retry on non-404 errors
            if verbose:
                print(f"Error deleting file: {e}")
            raise e


@retry_google_api(max_attempts=5, max_wait=60)
def copy_drive_file(drive_service, file_id, verbose=True):
    if verbose:
        print(f"Copying file with ID: {file_id}")

    # Setting name of copy to "tmp"
    request_body = {"name": "tmp"}
    drive_response = drive_service.files().copy(fileId=file_id, body=request_body, supportsAllDrives=True).execute()
    copy_document_id = drive_response.get("id")
    if verbose:
        print(f"Created a copy of the document with ID: {copy_document_id}")
    return copy_document_id


@retry_google_api(max_attempts=5, max_wait=60)
def update_google_doc(docs_service, document_id, form_data):
    # Retrieve the documents contents from the Docs service.
    docs_service.documents().get(documentId=document_id).execute()

    # Update the document content
    requests = []
    for key, value in form_data.items():
        requests.append({"replaceAllText": {"containsText": {"text": key, "matchCase": "true"}, "replaceText": value}})

    # Execute the update
    docs_service.documents().batchUpdate(documentId=document_id, body={"requests": requests}).execute()


@retry_google_api(max_attempts=5, max_wait=60)
def export_google_doc_as_pdf(drive_service, document_id, output_pdf_path):
    # Export the document as PDF
    request = drive_service.files().export_media(fileId=document_id, mimeType="application/pdf")
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print(f"Download {int(status.progress() * 100)}%.")

    # Save the PDF
    fh.seek(0)
    with open(output_pdf_path, "wb") as f:
        f.write(fh.getvalue())


@retry_google_api(max_attempts=5, max_wait=60)
def get_google_doc_content(docs_service, document_id):
    document = docs_service.documents().get(documentId=document_id).execute()
    content = document.get("body").get("content", [])
    return content
