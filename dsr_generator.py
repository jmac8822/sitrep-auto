import os
import shutil
from datetime import datetime
from google_drive_handler import download_latest_files
from dsr_formatter import generate_dsr_docx
from email_util import send_email_with_attachment

# === Config ===
SENDER_EMAIL = 'johnmhughes@gmail.com'  # Replace with your sender email
SENDER_PASSWORD = 'izcvszicmrqsodfh' # Use Gmail app password (not raw password)
RECIPIENTS = ['johnmhughes@gmail.com']

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

def main():
    print("📥 Downloading latest DSR and Tracker files from Google Drive...")
    dsr_path, tracker_path = download_latest_files()

    print("📊 Parsing Excel data and formatting DSR...")
    parsed_data = extract_dummy_data_from_excel(dsr_path)  # Placeholder

    ship_data = parsed_data['ships'][0]
    ship_name = ship_data['name'].replace(' ', '_')
    today_str = datetime.now().strftime("%Y-%m-%d")

    # Output file
    output_file = f"downloads/{ship_name}_{today_str}_DSR_Report.docx"
    temp_output_path = os.path.join('downloads', output_filename)

    # Generate doc
    generate_dsr_docx(parsed_data, temp_output_path)
    print(f"✅ DSR generated: {temp_output_path}")

    # 💾 Move to Gen_DSR/<ShipName>/ folder
    main_dir = "Gen_DSR"
#    ship_folder = os.path.join(main_dir, ship_name)
    ship_folder = os.path.join("Gen_DSR", ship_name)
    os.makedirs(ship_folder, exist_ok=True)

    final_path = os.path.join(ship_folder, output_filename)
    shutil.move(temp_output_path, final_path)
    print(f"📂 Moved DSR to: {final_path}")

    # 📧 Email it out
    subject = f"Daily SITREP – {today_str}"
    body = f"Attached is the daily DSR for {ship_data['name']} as of {today_str}."
    print(f"📧 Sending email to {RECIPIENTS}...")
    send_email_with_attachment(subject, body, final_path, RECIPIENTS, SENDER_EMAIL, SENDER_PASSWORD)

if __name__ == '__main__':
    main()
