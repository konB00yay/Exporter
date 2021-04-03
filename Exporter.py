from __future__ import print_function
from PyPDF2 import PdfFileWriter
from wand.image import Image
import PyPDF2
import os
import sys
import os.path
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

SCOPES = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]

APPLICATION_PATH = ''


def authorize():
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    elif __file__:
        application_path = os.path.dirname(__file__)
    global APPLICATION_PATH
    APPLICATION_PATH = application_path
    credentials = None

    cred_json_path = os.path.join(application_path, 'credentials.json')
    token_path = os.path.join(application_path, 'token.json')

    if os.path.exists(token_path):
        credentials = Credentials.from_authorized_user_file(token_path, SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not credentials or not credentials.valid:
        if credentials and credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                cred_json_path, SCOPES)
            credentials = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(token_path, 'w') as token:
            token.write(credentials.to_json())

    return build('drive', 'v3', credentials=credentials)


def main(drive_service):
    spreadsheet_id = input("Enter the sheet ID: ")

    sheet_metadata = drive_service.files().get(fileId=spreadsheet_id).execute()
    title = sheet_metadata.get('name')
    product_name = title.replace(" ", "_")
    filename = os.path.join(APPLICATION_PATH, product_name + ".pdf")
    original_pdf = filename

    request = drive_service.files().export(fileId=spreadsheet_id, mimeType='application/pdf')
    fh = open(filename, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print("Download %d%%." % int(status.progress() * 100))

    print('PDF exported: ' + filename)
    reader = PyPDF2.PdfFileReader(open(filename, 'rb'))
    page0 = reader.getPage(0)
    h = page0.mediaBox.getHeight()
    w = page0.mediaBox.getWidth()

    num_of_pages = reader.getNumPages()

    new_height = h * num_of_pages

    if num_of_pages > 1:
        new_pdf_page = PyPDF2.pdf.PageObject.createBlankPage(None, w, new_height)
        for i in range(num_of_pages):
            next_page = reader.getPage(i)
            if i != num_of_pages-1:
                new_pdf_page.mergeScaledTranslatedPage(next_page, 1, 0, h * (num_of_pages - i - 1))
            else:
                new_pdf_page.mergeScaledTranslatedPage(next_page, 1, 0, 120)
        writer = PdfFileWriter()
        writer.addPage(new_pdf_page)

        filename = title.replace(" ", "_") + "_resized" + ".pdf"

        with open(filename, 'wb') as f:
            writer.write(f)

        print('PDF resized for multiple pages: ' + filename)

    with(Image(filename=filename, resolution=500)) as source:
        for i, image in enumerate(source.sequence):
            left_coord = round(2048 - source.width/2)
            background = Image(height=4096, width=4096, resolution=500, background="#ffffff")
            background.composite(image, left_coord, 0, 'atop')
            background.compression_quality = 99
            background.transform(resize='2048x2048>')
            high_res = Image(background, resolution=800)
            new_filename = os.path.join(APPLICATION_PATH, product_name + ".jpg")
            high_res.save(filename=new_filename)

    print('JPG created: ' + new_filename)

    if os.path.exists(filename):
        os.remove(filename)
        print('Removed resized PDF')
    if os.path.exists(original_pdf):
        os.remove(original_pdf)
        print('Removed original PDF')


if __name__ == '__main__':
    service = authorize()
    main(drive_service=service)
