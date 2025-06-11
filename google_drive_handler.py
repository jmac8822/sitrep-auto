
import os
import io
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2 import service_account

SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
SERVICE_ACCOUNT_FILE = 'credentials.json'
MAIN_FOLDER_ID = '1Zr7FN_lWh6CIGxHXLNAzOyFzKJL5-bB3'

def create_drive_service():
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    service = build('drive', 'v3', credentials=credentials)
    return service

def list_all_files(service, folder_id):
    files = []
    queue = [folder_id]

    while queue:
        current_folder = queue.pop()
        results = service.files().list(
            q=f"'{current_folder}' in parents and trashed = false",
            fields="files(id, name, mimeType, modifiedTime, parents)").execute()
        items = results.get('files', [])
        for item in items:
            if item['mimeType'] == 'application/vnd.google-apps.folder':
                queue.append(item['id'])
            else:
                files.append(item)
    return files

def download_file(service, file_id, file_name):
    request = service.files().get_media(fileId=file_id)
    fh = io.FileIO(file_name, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return os.path.abspath(file_name)

def download_latest_files():
    service = create_drive_service()
    all_files = list_all_files(service, MAIN_FOLDER_ID)

    dsr_files = [f for f in all_files if f['name'].lower().endswith('.xlsx') and 'dsr' in f['name'].lower()]
    tracker_files = [f for f in all_files if 'master itsis installation tracker' in f['name'].lower()]

    if not dsr_files:
        raise Exception("No DSR Excel file found.")
    if not tracker_files:
        raise Exception("No tracker Excel file found.")

    latest_dsr = max(dsr_files, key=lambda f: f['modifiedTime'])
    latest_tracker = max(tracker_files, key=lambda f: f['modifiedTime'])

    print(f"📄 Latest DSR file: {latest_dsr['name']}")
    print(f"📄 Latest Tracker file: {latest_tracker['name']}")

    dsr_path = download_file(service, latest_dsr['id'], latest_dsr['name'])
    tracker_path = download_file(service, latest_tracker['id'], latest_tracker['name'])

    return dsr_path, tracker_path
