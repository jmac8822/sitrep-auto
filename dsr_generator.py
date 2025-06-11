import os
import shutil
from datetime import datetime
from google_drive_handler import download_latest_files, create_drive_service, upload_to_drive
from dsr_formatter import generate_dsr_docx
from email_util import send_email_with_attachment
import re

# === Config ===
SENDER_EMAIL = os.environ.get('EMAIL_USER')
SENDER_PASSWORD = os.environ.get('EMAIL_PASS')
RECIPIENTS = [os.environ.get('RECIPIENT_EMAIL')]

def extract_dummy_data_from_excel(excel_path):
    # Replace this with real Excel parsing logic
    return {
        'ships': [
            {
                'name': 'USS HOWARD',
                'region': 'Japan',
                'percent_t52a': '47%',
                'work_completed_t52a': [
                    'Cables Installed:',
                    '– R-PB111-W44086',
                    '– R-PB111-W44087 – 50% complete',
                    '– R-PB111-W44088 – 50% complete'
                ],
                'planned_work_t52a': ['Continue cable installation'],
                'percent_ecs': '2%',
                'work_completed_ecs': ['No production activity performed'],
                'next_steps_ecs': ['Awaiting ILS and workbook to initiate check-in']
            }
        ]
    }

def extract_full_ship_name(filename):
    match = re.search(r"USS [A-Z\- ]+\(DDG-\d+\)", filename)
    return match.group(0) if match else "Unknown_Ship"

def main():
    print("📥 Downloading latest DSR and Tracker files from Google Drive...")
    dsr_path, tracker_path = download_latest_files()

    print("📊 Parsing Excel data and formatting DSR...")
    parsed_data = extract_dummy_data_from_excel(dsr_path)

    dsr_filename = os.path.basename(dsr_path)
    ship_name_raw = extract_full_ship_name(dsr_filename)
    ship_folder_name = ship_name_raw.replace(' ', '_').replace('(', '').replace(')', '')
    today_str = datetime.now().strftime("%Y-%m-%d")

    output_filename = f"{ship_folder_name}_{today_str}_DSR_Report.docx"
    temp_output_path = os.path.join("downloads", output_filename)

    generate_dsr_docx(parsed_data, temp_output_path)
    print(f"✅ DSR generated: {temp_output_path}")

    ship_folder = os.path.join("Gen_DSR", ship_folder_name)
    os.makedirs(ship_folder, exist_ok=True)
    final_path = os.path.join(ship_folder, output_filename)
    shutil.move(temp_output_path, final_path)
    print(f"📂 Moved DSR to: {final_path}")

    subject = f"Daily SITREP – {today_str}"
    body = f"Attached is the daily DSR for {ship_name_raw} as of {today_str}."
    print(f"📧 Sending email to {RECIPIENTS}...")
    send_email_with_attachment(subject, body, final_path, RECIPIENTS, SENDER_EMAIL, SENDER_PASSWORD)

    service = create_drive_service()
    drive_file_id = upload_to_drive(service, final_path, ship_folder_name)
    print(f"📤 Uploaded to Google Drive in folder '{ship_folder_name}' (File ID: {drive_file_id})")

if __name__ == '__main__':
    main()
