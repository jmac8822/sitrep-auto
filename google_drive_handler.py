from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2 import service_account
import io
import os
from datetime import datetime

FOLDER_ID = '1Zr7FN_lWh6CIGxHXLNAzOyFzKJL5-bB3'
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
SERVICE_ACCOUNT_FILE = 'credentials.json'

def create_drive_service():
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    return build('drive', 'v3', credentials=credentials)

def list_files_in_folder(service, folder_id, name_contains=None):
    query = f"'{folder_id}' in parents and trashed = false"
    if name_contains:
        query += f" and name contains '{name_contains}'"

    results = service.files().list(
        q=query,
        fields="files(id, name, createdTime, mimeType)"
    ).execute()

    return results.get('files', [])

def download_file(service, file_id, file_name, destination_folder):
    request = service.files().get_media(fileId=file_id)
    fh = io.FileIO(os.path.join(destination_folder, file_name), 'wb')
    downloader = MediaIoBaseDownload(fh, request)

    done = False
    while not done:
        status, done = downloader.next_chunk()
        print(f"Downloading {file_name}: {int(status.progress() * 100)}%")

    return os.path.join(destination_folder, file_name)

def download_latest_files():
    service = create_drive_service()
    os.makedirs('downloads', exist_ok=True)

    subfolders = list_files_in_folder(service, FOLDER_ID)
    excel_file = None
    latest_time = None

    for folder in subfolders:
        files = list_files_in_folder(service, folder['id'], name_contains='DSR')
        for file in files:
            if file['mimeType'].endswith('spreadsheet') or file['name'].endswith('.xlsx'):
                created = datetime.fromisoformat(file['createdTime'].replace('Z', '+00:00'))
                if latest_time is None or created > latest_time:
                    latest_time = created
                    excel_file = file

    if not excel_file:
        raise Exception("No DSR Excel file found.")

    downloaded_path = download_file(service, excel_file['id'], excel_file['name'], 'downloads')

    tracker_files = list_files_in_folder(service, FOLDER_ID, name_contains='MASTER ITSIS')
    tracker_files.sort(key=lambda f: f['createdTime'], reverse=True)

    tracker_path = None
    if tracker_files:
        tracker = tracker_files[0]
        tracker_path = download_file(service, tracker['id'], tracker['name'], 'downloads')

    return downloaded_path, tracker_path
