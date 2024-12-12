import csv
import os
from Utilities.utils import get_files, get_single_file, initialize_necessary_files, check_gmail_app_password, check_body_template, check_csv, sort_csv


def extract_spreadsheet(spreadsheet_file_path, tosend_csv_path, wordlist_file_path):
    # Process the file
    with open(spreadsheet_file_path, mode='r', encoding="utf-8") as sheet_csv_file:
        reader = csv.DictReader(sheet_csv_file)
        try:
            # Create and write to 'tosend.csv'
            with open(tosend_csv_path, mode='w', newline='') as tosend_csv_file:
                csv_writer = csv.writer(tosend_csv_file)
                csv_writer.writerow(['Full Name', 'Email'])  # Write header row
                
                # Create and write to 'wordlist.txt'
                with open(wordlist_file_path, mode='w') as wordlist_file:
                    for row in reader:
                        if row['Attendance'].strip().upper() == 'TRUE':  # Check Attendance
                            full_name = row['Full Name'].strip().title()
                            csv_writer.writerow([full_name, row['Email']])
                            wordlist_file.write(f"{full_name}\n")
                    print("\n\'Full Name\' column successfully written to wordlist file.")
                    
                print("\'Full Name\' and \'Email\' columns successfully extracted to \'tosend.csv\' file.")
            print("\nSuccessfully extracted the spreadsheet file.\n")
        except PermissionError:
            print("\nFailed to write to \"tosend.csv\" file.\nEnsure that the file is not open on the system.\n")
            exit(1)


# === MAIN ENTRY POINT ===
if __name__ == "__main__":
    
    if not os.getcwd()[-9:] == "Resources":
        print("\nPlease change your working directory to the main repository.\n\nExiting...\n")
        exit(1)
        
    EMAIL_SUBJECT = "Subject"

    CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH = os.path.join(os.getcwd(),"Certificate_Email_Automation")
    CERTIFICATE_GENERATOR_DIRECTORY_PATH = os.path.join(os.getcwd(), "Certificate_Generator")
    EMAIL_SENDER_DIRECTORY_PATH = os.path.join(os.getcwd(), "Email_Sender")
    SPREADSHEET_DIRECTORY_PATH = os.path.join(CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH, "Spreadsheet")
    WORDLIST_DIRECTORY_PATH = os.path.join(CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH, "Wordlist")
    TEMPLATE_DIRECTORY_PATH = os.path.join(CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH, 'Certificate_Template')
    
    BODY_TEMPLATE_FILE_PATH = os.path.join(CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH, "body_template.html")
    LOG_FILE_PATH = os.path.join(CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH, "email_log.txt")
    GMAIL_APP_PASSWORD_FILE_PATH  = os.path.join(CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH, "gmail_app_password.txt")
    
    print("\n" + " Certificate_Email_Automation ".center(40, "-"))
    
    os.makedirs(CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH, exist_ok=True)
    os.makedirs(SPREADSHEET_DIRECTORY_PATH, exist_ok=True)
    os.makedirs(TEMPLATE_DIRECTORY_PATH, exist_ok=True)
    
    initialize_necessary_files(BODY_TEMPLATE_FILE_PATH, None, GMAIL_APP_PASSWORD_FILE_PATH)
    
    csv_files = get_files(SPREADSHEET_DIRECTORY_PATH, 'CSV')
    spreadsheet_file = get_single_file('Spreadsheet', SPREADSHEET_DIRECTORY_PATH, 'CSV')
    spreadsheet_file_path = os.path.join(SPREADSHEET_DIRECTORY_PATH, spreadsheet_file)
    
    tosend_csv_path = os.path.join(CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH, "tosend.csv")
    certificate_script_path = os.path.join(CERTIFICATE_GENERATOR_DIRECTORY_PATH, "certificate_generator.py")
    email_script_path = os.path.join(EMAIL_SENDER_DIRECTORY_PATH, "send_email.py")
    
    passwd = check_gmail_app_password(GMAIL_APP_PASSWORD_FILE_PATH)
    
    os.makedirs(WORDLIST_DIRECTORY_PATH, exist_ok=True)
    text_files = get_files(WORDLIST_DIRECTORY_PATH, 'TXT')
    if not text_files:
        with open(os.path.join(WORDLIST_DIRECTORY_PATH, 'wordlist.txt'), 'w') as tosend_csv_file:
            pass
        
    wordlist_file_path = os.path.join(WORDLIST_DIRECTORY_PATH, 'wordlist.txt')
    
    check_body_template(BODY_TEMPLATE_FILE_PATH)
    
    # Open the file and ensure it has the correct contents as needed
    check_csv(spreadsheet_file_path, "Other", "Attendance")
    
    extract_spreadsheet(spreadsheet_file_path, tosend_csv_path, wordlist_file_path)

    sort_csv(spreadsheet_file_path)
    sort_csv(tosend_csv_path)
    
    try:
        certificate_script_status = os.system(f"python {certificate_script_path} extract_certify_and_email_script")
        if certificate_script_status == 0:
            os.system(f"python {email_script_path} extract_certify_and_email_script \"{EMAIL_SUBJECT}\"")
    except KeyboardInterrupt:
        exit(1)
    except Exception as e:
        print(f"\n\nAn error occured while executing the script.\n{e}\n")
        exit(1)
