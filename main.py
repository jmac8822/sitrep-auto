
import os
import datetime
from google_drive_handler import download_latest_files
from dsr_generator import generate_dsr
from email_sender import send_email
from utils import has_data_changed, update_last_processed

def main():
    today = datetime.datetime.now().date()
    if not has_data_changed(today):
        print("No new data detected. Skipping DSR generation.")
        return

    latest_dsr_file = download_latest_files()
    if latest_dsr_file:
        report_path = generate_dsr(latest_dsr_file)
        send_email(report_path)
        update_last_processed(today)
        print("DSR generated and emailed successfully.")
    else:
        print("No DSR input file found.")

if __name__ == "__main__":
    main()
