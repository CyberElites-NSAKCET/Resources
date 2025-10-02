import os
import csv
import sys

# Get the parent directory, add it to python path and import the modules
parent_dir = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
sys.path.append(parent_dir)

from Utilities.utils import check_attendance, check_body_template, check_csv, check_gmail_app_password, clean_csv_fieldnames, get_files, get_single_file, initialize_necessary_files, load_config, sort_csv


## ===========================================================================
### Functions

# Functiont to extract spreadsheet and write necessary columns to wordlist and csv file
def extract_spreadsheet(spreadsheet_file_path, tosend_csv_path, wordlist_file_path):
    """
    Extracts specific columns from a spreadsheet file and creates a filtered CSV file
    and a wordlist text file for further processing.

    Args:
        spreadsheet_file_path (str): Path to the input spreadsheet CSV file.
                                     The file must have "Full Name", "Email", and "Attendance" columns.
        tosend_csv_path (str): Path to save the output CSV file containing "Full Name" and "Email" columns.
        wordlist_file_path (str): Path to save the text file containing "Full Name" entries.

    Workflow:
        - Reads the spreadsheet file.
        - Filters rows where the "Attendance" column is marked as "TRUE" (case-insensitive).
        - Extracts "Full Name" and "Email" columns and writes them to the output CSV file.
        - Writes "Full Name" values to the wordlist text file.

    Raises:
        PermissionError: If the output CSV file is open or locked by another process.
        KeyError: If required columns ("Full Name", "Email", or "Attendance") are missing in the spreadsheet.
        Exception: For any other unforeseen errors.

    Prints:
        - Success messages upon successful file creation.
        - Error messages if file writing fails.

    """

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
                            csv_writer.writerow([full_name, row['Email'].strip()])
                            wordlist_file.write(f"{full_name}\n")
                    print("\n\'Full Name\' column successfully written to 'Wordlist\\wordlist.txt' file.")

                print("\'Full Name\' and \'Email\' columns successfully extracted to \'tosend.csv\' file.")
        except PermissionError:
            print("\nFailed to write to \"tosend.csv\" file.\nEnsure that the file is not open on the system.\n")
            sys.exit(1)


## ===========================================================================
# === MAIN ENTRY POINT ===

if __name__ == "__main__":
    """
    Certificate Email Automation Script

    This script automates the process of:
    1. Extracting recipient details from a spreadsheet.
    2. Generating certificates for attendees.
    3. Sending certificates via email with personalized messages.

    Workflow:
        - Ensures necessary directories and files are created and initialized.
        - Reads the spreadsheet file to extract attendee details based on "Attendance".
        - Generates certificates for attendees marked as present.
        - Sends emails with the generated certificates as attachments.

    Environment Setup:
        - `Certificate_Email_Automation` directory contains all necessary files, including:
            - `Spreadsheet`: Directory for input CSV files containing attendee details.
            - `Certificate_Template`: Directory for certificate templates.
            - `Wordlist`: Directory for wordlists used in certificate generation.

    Key Components:
        - **Spreadsheet Extraction**:
            Extracts "Full Name" and "Email" columns for attendees marked as present.
        - **Certificate Generation**:
            Calls `certificate_generator.py` to create certificates.
        - **Email Sending**:
            Calls `send_email.py` to send personalized emails with certificates.

    Raises:
        - KeyError: If required columns are missing in the spreadsheet.
        - FileNotFoundError: If the input spreadsheet file is not found.
        - PermissionError: If the output files are locked or inaccessible.
        - Exception: For other unforeseen errors during execution.

    Dependencies:
        - `Utilities.utils` module for common file handling and validation functions.
        - External scripts: `certificate_generator.py` and `send_email.py`.

    Logs:
        - Actions and errors are logged in the `email_log.txt` file for review.
    """

    CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH = os.path.abspath(os.path.dirname(__file__))
    ROOT_REPO_PATH = os.path.abspath(os.path.dirname(CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH))

    CERTIFICATE_GENERATOR_DIR_PATH = os.path.join(ROOT_REPO_PATH, "Certificate_Generator")
    EMAIL_SENDER_DIR_PATH = os.path.join(ROOT_REPO_PATH, "Email_Sender")

    CERTIFICATE_TEMPLATE_DIR_PATH = os.path.join(CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH, 'Certificate_Template')
    WORDLIST_DIR_PATH = os.path.join(CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH, "Wordlist")
    SPREADSHEET_DIR_PATH = os.path.join(CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH, "Spreadsheet")
    BODY_TEMPLATE_FILE_PATH = os.path.join(CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH, "cert-email_html_body_template.html")
    LOG_FILE_PATH = os.path.join(CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH, "cert-email_log.txt")

    print("\n" + " Certificate_Email_Automation ".center(40, "-"))

    os.makedirs(CERTIFICATE_TEMPLATE_DIR_PATH, exist_ok=True)
    os.makedirs(SPREADSHEET_DIR_PATH, exist_ok=True)

    config = load_config(os.path.join(CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH, "cert-email_config.json"))
    initialize_necessary_files(BODY_TEMPLATE_FILE_PATH)

    sender_email = config.get("sender_email", "").strip()
    email_subject = config.get("email_subject", "").strip()
    passwd = config.get("gmail_app_password", "").strip()

    if sender_email == "" or email_subject in ["", "Subject"]:
        print("\nError: Please configure all fields properly in the config file before running the script.\n\nExiting...\n")
        sys.exit(1)
    check_gmail_app_password(passwd)

    csv_files = get_files(SPREADSHEET_DIR_PATH, 'CSV')
    spreadsheet_file = get_single_file('Spreadsheet', SPREADSHEET_DIR_PATH, 'CSV')
    spreadsheet_file_path = os.path.join(SPREADSHEET_DIR_PATH, spreadsheet_file)

    tosend_csv_path = os.path.join(CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH, "tosend.csv")
    certificate_script_path = os.path.join(CERTIFICATE_GENERATOR_DIR_PATH, "certificate_generator.py")
    email_script_path = os.path.join(EMAIL_SENDER_DIR_PATH, "send_email.py")

    os.makedirs(WORDLIST_DIR_PATH, exist_ok=True)
    text_files = get_files(WORDLIST_DIR_PATH, 'TXT')
    if not text_files:
        with open(os.path.join(WORDLIST_DIR_PATH, 'wordlist.txt'), 'w') as tosend_csv_file:
            pass

    wordlist_file_path = os.path.join(WORDLIST_DIR_PATH, 'wordlist.txt')

    check_body_template(BODY_TEMPLATE_FILE_PATH)

    clean_csv_fieldnames(spreadsheet_file_path)

    # Open the file and ensure it has the correct contents as needed
    check_csv(spreadsheet_file_path, "Other", "Attendance")

    check_attendance(spreadsheet_file_path)

    sort_csv(spreadsheet_file_path)

    extract_spreadsheet(spreadsheet_file_path, tosend_csv_path, wordlist_file_path)

    print("\nSorting extracted 'tosend.csv' file contents:",end='')
    sort_csv(tosend_csv_path)

    try:
        certificate_script_status = os.system(f'python "{certificate_script_path}" extract_certify_and_email_script')
        if certificate_script_status == 0:
            os.system(f'python "{email_script_path}" extract_certify_and_email_script')
    except (KeyboardInterrupt, EOFError):
        sys.exit(1)
    except Exception as e:
        print(f"\n\nAn error occured while executing the script.\n{e}\n")
        sys.exit(1)
