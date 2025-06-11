def main():
    print("📥 Downloading latest DSR and Tracker files from Google Drive...")
    dsr_path, tracker_path = download_latest_files()

    print("📊 Parsing Excel data and formatting DSR...")
    parsed_data = extract_dummy_data_from_excel(dsr_path)  # Placeholder

    ship_data = parsed_data['ships'][0]
    ship_name = ship_data['name'].replace(' ', '_')
    today_str = datetime.now().strftime("%Y-%m-%d")

    output_filename = f"{ship_name}_{today_str}_DSR_Report.docx"
    temp_output_path = os.path.join("downloads", output_filename)

    # Generate the DSR doc
    generate_dsr_docx(parsed_data, temp_output_path)
    print(f"✅ DSR generated: {temp_output_path}")

    # Move to Gen_DSR/<ShipName>/
    ship_folder = os.path.join("Gen_DSR", ship_name)
    os.makedirs(ship_folder, exist_ok=True)
    final_path = os.path.join(ship_folder, output_filename)
    shutil.move(temp_output_path, final_path)
    print(f"📂 Moved DSR to: {final_path}")

    # Send email
    subject = f"Daily SITREP – {today_str}"
    body = f"Attached is the daily DSR for {ship_data['name']} as of {today_str}."
    print(f"📧 Sending email to {RECIPIENTS}...")
    send_email_with_attachment(subject, body, final_path, RECIPIENTS, SENDER_EMAIL, SENDER_PASSWORD)

    # ✅ Upload to Google Drive
    from google_drive_handler import create_drive_service, upload_to_drive
    service = create_drive_service()
    drive_file_id = upload_to_drive(service, final_path, ship_name)
    print(f"📤 Uploaded to Google Drive in folder '{ship_name}' (File ID: {drive_file_id})")

if __name__ == '__main__':
    main()
