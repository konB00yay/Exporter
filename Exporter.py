from __future__ import print_function
import os
from PyPDF2 import PdfFileWriter
import PyPDF2
import sys
import subprocess
import os.path
import re
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

SCOPES = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]

global APPLICATION_PATH
APPLICATION_PATH = ''


def throw_the_kitchen_sink():
    if os.system("which brew") != 0:
        c = input('Install Homebrew y/N: ')
        if c == 'y':
            command = f'/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/master/install.sh)"'
            print('Going to install Homebrew, this may take some time')
            i = subprocess.Popen(command, shell=True)
            i.wait()
        if os.system("which brew") != 0:
            print('Homebrew did not install')

        if os.system("which gs") != 0:
            gs = subprocess.Popen(['brew', 'install', 'ghostscript'])
            gs.wait()

        if os.system("which pkg-config") != 0:
            pk = subprocess.Popen(['brew', 'install', 'pkg-config'])
            pk.wait()

        if os.system("which convert") != 0:
            im = subprocess.Popen(['brew', 'install', 'imagemagick'])
            im.wait()

        if os.system("which gs") == 0:
            print('Installed gs')
        if os.system("which convert") != 0:
            username = subprocess.check_output("whoami").strip()
            new_command = 'sudo chown -R ' + username + ':admin /usr/local/include'
            i = subprocess.Popen(new_command, shell=True)
            i.wait()

            if os.system("which convert") == 0:
                print('imagemagick installed ready to go')

    else:
        if os.system("which gs") != 0:
            gs = subprocess.Popen(['brew', 'install', 'ghostscript'])
            gs.wait()

        if os.system("which pkg-config") != 0:
            pk = subprocess.Popen(['brew', 'install', 'pkg-config'])
            pk.wait()

        if os.system("which convert") != 0:
            im = subprocess.Popen(['brew', 'install', 'imagemagick'])
            im.wait()

        if os.system("which gs") == 0:
            print('Installed gs')
        if os.system("which convert") != 0:
            username = subprocess.check_output("whoami").strip()
            new_command = 'sudo chown -R ' + username + ':admin /usr/local/include'
            i = subprocess.Popen(new_command, shell=True)
            i.wait()

            if os.system("which convert") == 0:
                print('imagemagick installed ready to go')


def authorize():
    credentials = None

    cred_json_path = os.path.join(APPLICATION_PATH, 'credentials.json')
    token_path = os.path.join(APPLICATION_PATH, 'token.json')

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


def id_or_url(spreadsheet_id):
    # https://docs.google.com/spreadsheets/d/16RWH9XBBwd8pRYZDSo9EontzdVPqxdGnwM5MnP6T48c/edit#gid=0
    if spreadsheet_id.startswith("https://docs.google.com/spreadsheets/d/"):
        id = spreadsheet_id[39: spreadsheet_id.find("/", 39)]

    else:
        id = spreadsheet_id

    if re.match("^([a-zA-Z0-9]|_|-)+$", id) is None:
        raise ValueError("url argument must be an alphanumeric id or a full URL")
    return id


def main(drive_service):
    try:
        from wand.image import Image
    except ImportError as error:
        print('No ImageMagick Likely, installing.............')
        throw_the_kitchen_sink()
        from wand.image import Image

    spreadsheet_id = input("Enter sheet ID or URL: ")

    id = id_or_url(spreadsheet_id)

    sheet_metadata = drive_service.files().get(fileId=id, supportsAllDrives=True).execute()
    title = sheet_metadata.get('name')
    product_name = title.replace(" ", "_")
    filename = os.path.join(APPLICATION_PATH, product_name + ".pdf")
    original_pdf = filename

    request = drive_service.files().export(fileId=id, mimeType='application/pdf')
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

    with open(filename, 'rb') as f:
        blob = f.read()
        with(Image(blob=blob, resolution=500)) as source:
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
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    elif __file__:
        application_path = os.path.dirname(__file__)
    APPLICATION_PATH = application_path

    service = authorize()
    while (True):
        main(drive_service=service)
        quit_input = input('Press Enter to convert another sheet, if not then press q: ')

        if quit_input == 'q':
            sys.exit()
