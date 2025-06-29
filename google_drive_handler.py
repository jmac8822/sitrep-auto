import os
import io
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.http import MediaIoBaseDownload

SCOPES = ['https://www.googleapis.com/auth/drive']
SERVICE_ACCOUNT_FILE = 'credentials.json'
DOWNLOAD_DIR = '/opt/render/project/src/downloads'
os.makedirs(DOWNLOAD_DIR, exist_ok=True)


def create_drive_service():
    credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    return build('drive', 'v3', credentials=credentials)


def download_file(service, file_id, file_name):
    if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
        request = service.files().get_media(fileId=file_id)
        file_path = os.path.join(DOWNLOAD_DIR, file_name)
    else:
        request = service.files().export_media(
            fileId=file_id,
            mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        file_path = os.path.join(DOWNLOAD_DIR, f"{file_name}.xlsx")

    fh = io.FileIO(file_path, 'wb')
    downloader = MediaIoBaseDownload(fh, request)

    done = False
    while not done:
        _, done = downloader.next_chunk()

    return file_path


def list_all_files(service, folder_id):
    files = []
    queue = [folder_id]

    while queue:
        current = queue.pop()
        result = service.files().list(
            q=f"'{current}' in parents and trashed = false",
            fields="files(id, name, mimeType, modifiedTime, parents)"
        ).execute()
        items = result.get('files', [])
        for item in items:
            if item['mimeType'] == 'application/vnd.google-apps.folder':
                queue.append(item['id'])
            else:
                files.append(item)
    return files


def download_latest_files():
    MAIN_FOLDER_ID = '1Zr7FN_lWh6CIGxHXLNAzOyFzKJL5-bB3'
    service = create_drive_service()
    all_files = list_all_files(service, MAIN_FOLDER_ID)

    # Ensure this checks ALL files
    import re

    dsr_files = [
        f for f in all_files
        if (
            f['name'].lower().endswith('.xlsx') and
            re.search(r'\bdsr\b', f['name'].lower()) and
            re.search(r'\d{4}-\d{2}-\d{2}', f['name'])  # has YYYY-MM-DD
        )
    ]

    tracker_files = [
        f for f in all_files
        if 'master itsis installation tracker' in f['name'].lower()
    ]

    if not dsr_files:
        raise Exception("No DSR Excel file found.")
    if not tracker_files:
        raise Exception("No tracker Excel file found.")

    # ✅ Select based on most recent modifiedTime
    latest_dsr = max(dsr_files, key=lambda f: f['modifiedTime'])
    latest_tracker = max(tracker_files, key=lambda f: f['modifiedTime'])

    print(f"📄 Latest DSR file: {latest_dsr['name']}")
    print(f"📄 Latest Tracker file: {latest_tracker['name']}")

    dsr_path = download_file(service, latest_dsr['id'], latest_dsr['name'])
    tracker_path = download_file(service, latest_tracker['id'], latest_tracker['name'])

    return dsr_path, tracker_path


def upload_to_drive(service, local_file_path, ship_name):
    # Main shared folder ID
    parent_id = '1GMZQgdsJ3fyhUvAX_mAyRNpuOIou7PGz'

    # Check or create subfolder for the ship
    query = f"'{parent_id}' in parents and name = '{ship_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    response = service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
    items = response.get('files', [])

    if items:
        folder_id = items[0]['id']
    else:
        file_metadata = {
            'name': ship_name,
            'mimeType': 'application/vnd.google-apps.folder',
            'parents': [parent_id]
        }
        folder = service.files().create(body=file_metadata, fields='id').execute()
        folder_id = folder['id']

    # ✅ Upload the file to the ship's subfolder
    from googleapiclient.http import MediaFileUpload
    file_metadata = {
        'name': os.path.basename(local_file_path),
        'parents': [folder_id]
    }
    media = MediaFileUpload(
        local_file_path,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )
    uploaded = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id'
    ).execute()

    return uploaded['id']
