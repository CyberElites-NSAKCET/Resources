import os
import sys

# Get the parent directory, add it to python path and import the modules
parent_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.append(parent_dir)

import csv
import logging
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from socket import gaierror
try:
    from Utilities.utils import add_attachment, check_attachments, check_body_template, check_csv, check_gmail_app_password, clean_csv_fieldnames, get_files, get_single_file, initialize_necessary_files, load_config, read_email_body_template, sort_csv
except ImportError:
    print("\nThis script requires the \'Utilities\' module.\n\nPlease ensure that the script is run from the correct directory.\n\nExiting...\n")
    exit(1)


## ===========================================================================
### Functions

# === FUNCTION: SEND EMAIL ===
def send_email(recipient_email, name, subject, body, attachments):
    """
    Sends an email to a single recipient with optional attachments.

    Args:
        recipient_email (str): The recipient's email address.
        name (str): The recipient's name for personalization.
        subject (str): The subject of the email.
        body (str): The HTML content of the email body.
        attachments (list or str): List of attachment file names/paths, or a single file name/path depending on attachment mode.

    Raises:
        smtplib.SMTPAuthenticationError: If authentication with the SMTP server fails.
        socket.gaierror: If there is a network connection issue.
        Exception: For other unforeseen errors during the email sending process.

    Logs:
        - Successful email delivery with recipient's email.
        - Errors encountered while sending the email.
    """

    try:
        # Set up the email
        msg = MIMEMultipart()
        msg["From"] = SENDER_EMAIL
        msg["To"] = recipient_email
        msg["Subject"] = subject

        # Add the HTML body
        msg.attach(MIMEText(body, "html"))

        # Check command-line arguments
        if len(sys.argv) > 1 and sys.argv[1].strip() == "extract_certify_and_email_script":
            gen_certs_dir_path = os.path.join(CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH, "gen_certs_dir_path.txt")

            with open(gen_certs_dir_path, "r") as file:
                gen_certs_dir_path = file.read()

            attachment_path = os.path.join(gen_certs_dir_path, attachments)
            add_attachment(msg, attachment_path)

        else:
            for attachment_path in attachments:
                if attachment_path.strip():  # Ensure path is not empty
                    attachment_path = os.path.join(ATTACHMENTS_DIRECTORY_PATH, attachment_path)

                    add_attachment(msg, attachment_path)
        try:
            # Connect to Gmail's SMTP server and send the email
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls()
                server.login(SENDER_EMAIL, SENDER_PASSWORD)
                server.sendmail(
                    SENDER_EMAIL,
                    [recipient_email],
                    msg.as_string(),
                )
        except KeyboardInterrupt:
            print(f"\nKeyboard Interrupt!\n\nEmails not sent form recipient name: \'{name}\'\n\nExiting...\n")
            exit(1)

        # Log success
        logging.info(f"Email sent to {recipient_email}")
        print(f"Email sent to {recipient_email}")

    except smtplib.SMTPAuthenticationError as e:
        print(f"Incorrect Gmail App Password!\nAuthentication Failed for \'{SENDER_EMAIL}\' with provided password.\n")
        exit(1)
    except gaierror as e:
        print("Failed to send Emails....\nCheck your Internet connection\nEmails not sent form recipient name: \'{name}\'\n\nExiting...\n")
        exit(1)
    except Exception as e:
        # Log failure
        logging.error(f"Failed to send email to {recipient_email}: {e}")
        print(f"Failed to send email to {recipient_email}: {e}")


## --------------------------------------------------------------------------
# === FUNCTION: SEND BULK EMAILS ===
def send_bulk_emails(csv_file_path, body_template_file):
    """
    Sends bulk emails to recipients by reading their details from a CSV file.

    Args:
        csv_file_path (str): Path to the CSV file containing recipient details.
                             The file should include columns for "Email" and "Full Name".
                             If attachments are needed, include an "Attachments" column.
        body_template_file (str): Path to the HTML file used as the email body template.
                                  The template should include placeholders for customization
                                  (e.g., "{{name}}").

    Raises:
        FileNotFoundError: If the specified CSV file does not exist.
        ValueError: If the CSV file format is invalid or lacks required fields.
        Exception: For other unforeseen errors during the email sending process.

    Logs:
        - Successful email delivery for each recipient.
        - Errors encountered while processing individual rows of the CSV file.
    """

    global ATTACHMENT_MODE

    try:
        # Read the email body template
        body_template = read_email_body_template(body_template_file)

        # Open the CSV file
        with open(csv_file_path, "r", encoding="utf-8") as csv_file:
            csv_file.seek(0)  # Reset file pointer
            reader = csv.DictReader(csv_file)

            # Read the common attachments if needed
            if ATTACHMENT_MODE == "Common":
                first_row = next(reader, None)
                if first_row:
                    common_attachments = (
                        first_row["Attachments"].split(";")
                        if first_row.get("Attachments")
                        else []
                    )

            csv_file.seek(0)  # Reset file pointer
            reader = csv.DictReader(csv_file)  # Reinitialize reader

            print("\nSending emails to recipients.....\n")

            row_index = 2
            for row in reader:
                try:
                    if not row.get("Email", "") or not row.get("Full Name", ""):
                        raise ValueError("Missing recipient email or name in a row.")
                    else:
                        # Extract recipient details
                        recipient_email = row.get("Email", "").lower().strip()
                        name = row.get("Full Name", "").title().strip()

                    # Determine attachments
                    if ATTACHMENT_MODE == "Respective":
                        if row.get("Attachments", ""):
                            attachments = (row.get("Attachments", "").split(";") if row.get("Attachments", "").strip() else [])
                        else:
                            attachments = []

                    elif ATTACHMENT_MODE == "Common":
                        attachments = common_attachments

                    elif ATTACHMENT_MODE == "Other":
                        attachments = f"{name.title().strip().replace(' ', '_')}_certificate.pdf"

                    elif ATTACHMENT_MODE == "None":
                        attachments = []

                    else:
                        print("\nInvalid Attachment Mode specified!\nPlease select among \'Respective\',\'Common\' or \'None\'.")
                        exit(1)

                    # Customize the email body
                    personalized_body = body_template.replace("{{name}}", name)
                    # personalized_body = body_template.replace("{{phone}}", phone)

                    # Send the email
                    send_email(recipient_email, name, EMAIL_SUBJECT, personalized_body, attachments)

                    row_index += 1

                except Exception as row_error:
                    logging.error(f"Error processing recipient row\n  Row Index- \'{row_index}\' : {row_error}")
                    print(f"\nError processing recipient row\n  Row Index- \'{row_index}\' : {row_error}\n")

    except FileNotFoundError as fnf_error:
        logging.error(f"CSV file not found: {csv_file_path} - {fnf_error}")
        print(f"CSV file not found: {csv_file_path} - {fnf_error}")
    except ValueError as value_error:
        logging.error(f"Invalid CSV file format: {value_error}")
        print(f"Invalid CSV file format: {value_error}")
    except Exception as e:
        logging.error(f"Unexpected error: {e}")
        print(f"Unexpected error: {e}")


## ===========================================================================
# === MAIN ENTRY POINT ===

if __name__ == "__main__":
    """
    Main entry point for the email sender script.

    Features:
        - Sends bulk emails using recipient details from a CSV file.
        - Supports customizable email body via an HTML template.
        - Handles attachments in multiple modes:
            - 'None': No attachments.
            - 'Common': The same attachments (from the first recipient) are sent to all recipients.
            - 'Respective': Attachments are specified per recipient in the CSV file.
            - 'Other': Used by the automation script for certificate distribution.

    Automation Mode:
        - Integrates with a certificate generation module for automated certificate emailing.
        - Reads configuration and certificate paths from specific directories.

    Environment:
        - Sets up directories for logs, attachments, and templates.
        - Requires a Gmail App Password for authentication.

    Setup:
        - Store Gmail App Password securely as specified in the config.
        - Prepare CSV files, attachments, and email body templates before running.

    Logging:
        - Logs all email operations and errors to a log file with recipient details.

    Raises:
        - FileNotFoundError: If required files (CSV, HTML, or attachments) are missing.
        - ValueError: If invalid data is found in the CSV file.
        - smtplib.SMTPAuthenticationError: If Gmail App Password is incorrect.
        - socket.gaierror: For network connectivity issues.
        - Exception: For other unforeseen issues during email operations.
    """

    print("\n" + " Email Sender ".center(24, "-"))

    if os.getcwd()[-9:] == "Resources":
        EMAIL_SENDER_DIRECTORY_PATH = os.path.join(os.getcwd(), 'Email_Sender')
        ROOT_REPO_PATH = os.getcwd()

    elif os.getcwd()[-12:] == "Email_Sender":
        EMAIL_SENDER_DIRECTORY_PATH = os.getcwd()
        ROOT_REPO_PATH = os.path.join(os.getcwd(), '..')

    else:
        print("\nPlease change your working directory to the main repository.\n\nExiting...\n")
        exit(1)

    config = load_config()
    # === CONFIGURATION ===
    SMTP_SERVER = config.get("smtp_server")
    SMTP_PORT = config.get("smtp_port")
    SENDER_EMAIL = config.get("sender_email")
    EMAIL_SUBJECT = config.get("email_subject")
    SENDER_PASSWORD = config.get("gmail_app_password")
    ATTACHMENT_MODE = config.get("attachment_mode")

    automation_script = len(sys.argv) > 1 and sys.argv[1] == "extract_certify_and_email_script"

    if not automation_script and ATTACHMENT_MODE == "Other":
        print("\nInvalid Attachment mode specified for the Email_Sender code.\nThis mode is explicitly for the automation script\n\nPlease select from \"None\", \"Common\" or \"Respective\"\n\nExiting..\n")
        exit(1)

    if automation_script:
        CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH = os.path.join(ROOT_REPO_PATH, "Certificate_Email_Automation")
        DIR_PATH = CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH
        CSV_FILE_PATH = os.path.join(CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH, "tosend.csv")
        ATTACHMENT_MODE = "Other"
    else:
        DIR_PATH = EMAIL_SENDER_DIRECTORY_PATH
        ATTACHMENTS_DIRECTORY_PATH = os.path.join(EMAIL_SENDER_DIRECTORY_PATH, "Attachments")
        SPREADSHEET_DIRECTORY_PATH = os.path.join(EMAIL_SENDER_DIRECTORY_PATH, "Spreadsheet")
        os.makedirs(ATTACHMENTS_DIRECTORY_PATH, exist_ok=True)
        os.makedirs(SPREADSHEET_DIRECTORY_PATH, exist_ok=True)

    BODY_TEMPLATE_FILE_PATH = os.path.join(DIR_PATH, "body_template.html")
    LOG_FILE_PATH = os.path.join(DIR_PATH, "email_log.txt")

    initialize_necessary_files(BODY_TEMPLATE_FILE_PATH) if not automation_script else initialize_necessary_files(log_file=LOG_FILE_PATH)

    if not automation_script:
        check_gmail_app_password(SENDER_PASSWORD)
        csv_files = get_files(SPREADSHEET_DIRECTORY_PATH, 'CSV')
        spreadsheet_file = get_single_file('Spreadsheet', SPREADSHEET_DIRECTORY_PATH, 'CSV')
        CSV_FILE_PATH = os.path.join(SPREADSHEET_DIRECTORY_PATH, spreadsheet_file)

    check_body_template(BODY_TEMPLATE_FILE_PATH)

    # Check command-line arguments
    if automation_script:
        check_attachments(CSV_FILE_PATH, attachment_mode=ATTACHMENT_MODE,automation_dir_path=CERTIFICATE_EMAIL_AUTOMATION_DIR_PATH)
    else:
        clean_csv_fieldnames(CSV_FILE_PATH)
        sort_csv(CSV_FILE_PATH)
        check_csv(CSV_FILE_PATH, ATTACHMENT_MODE)
        check_attachments(CSV_FILE_PATH, ATTACHMENTS_DIRECTORY_PATH, ATTACHMENT_MODE)
        initialize_necessary_files(log_file=LOG_FILE_PATH)

    # === SET UP LOGGING ===
    logging.basicConfig(
        filename=LOG_FILE_PATH,
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )

    send_bulk_emails(CSV_FILE_PATH, BODY_TEMPLATE_FILE_PATH)
