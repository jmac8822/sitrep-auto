from google_drive_handler import download_latest_files
from dsr_formatter import generate_dsr_docx
from email_util import send_email_with_attachment
import os
from datetime import datetime

# === Config ===
SENDER_EMAIL = 'your_email@gmail.com'  # Replace with your sender email
SENDER_PASSWORD = 'your_app_password'  # Use Gmail app password (never raw password)
RECIPIENTS = ['johnmhughes@gmail.com']

def extract_dummy_data_from_excel(excel_path):
    # Replace with actual parsing logic — this is a placeholder structure
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

def main():
    print("📥 Downloading latest DSR and Tracker files from Google Drive...")
    dsr_path, tracker_path = download_latest_files()

    print("📊 Parsing Excel data and formatting DSR...")
    parsed_data = extract_dummy_data_from_excel(dsr_path)  # Placeholder logic

    today_str = datetime.now().strftime("%Y-%m-%d")
    ship_name = parsed_data['ships'][0]['name'].replace(' ', '_')
    output_file = f"downloads/USS_{ship_name}_{today_str}_DSR_Report.docx"

    generate_dsr_docx(parsed_data, output_file)
    print(f"✅ DSR generated: {output_file}")

    subject = f"Daily SITREP – {today_str}"
    body = f"Attached is the daily DSR for {parsed_data['ships'][0]['name']} as of {today_str}."
    print(f"📧 Sending email to {RECIPIENTS}...")
    send_email_with_attachment(subject, body, output_file, RECIPIENTS, SENDER_EMAIL, SENDER_PASSWORD)

if __name__ == '__main__':
    main()
