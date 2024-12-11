import csv
import logging
import os
import smtplib
import sys
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from socket import gaierror

# Get the parent directory, add it to python path and import the modules
parent_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
sys.path.append(parent_dir)
from Certificate_Generator.certificate_generator import get_files, get_single_file


# === FUNCTION: READ EMAIL BODY TEMPLATE ===
def read_email_body_template():
    """
    Reads the HTML email body content from a file.
    """
    try:
        with open(BODY_TEMPLATE_FILE_PATH, "r", encoding="utf-8") as file:
            return file.read()
    except FileNotFoundError:
        logging.error(f"Body template file not found: {BODY_TEMPLATE_FILE_PATH}")
        print(f"Body template file not found: {BODY_TEMPLATE_FILE_PATH}")
        return ""



def add_attachment(msg, attachment_path):
    try:
        attachment = MIMEBase("application", "octet-stream")
        with open(attachment_path.strip(), "rb") as file:
            attachment.set_payload(file.read())
        encoders.encode_base64(attachment)
        attachment.add_header(
            "Content-Disposition",
            f"attachment; filename={os.path.basename(attachment_path)}",
        )
        msg.attach(attachment)
    except FileNotFoundError:
        logging.error(f"Attachment not found: {attachment_path}")
        print(f"Attachment not found: {attachment_path}")



## --------------------------------------------------------------------------
# Function to sort the provided wordlist file
def sort_csv(file_path):
    """
    Sorts the csv file contents.

    Args:
        file_path (str): The path to the csv file.

    Returns:
        list: A list of sorted lines of the file
    """
    # Read, sort, and overwrite the file
    try:
        with open(file_path, 'r', newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            # Sort rows by the "Full Name" column
            sorted_rows = sorted(reader, key=lambda row: row["Full Name"])
            fieldnames = reader.fieldnames  # Preserve the original header order
    except Exception as e:
        print(f"\nError in reading TXT wordlist!\nPlease ensure that the file is not corrupted.\n\nExiting...\n")
        exit(1)

    try:
        # Overwrite the same file with sorted data
        with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(sorted_rows)

    except Exception as e:
        print(f"\nError sorting the csv file.\nMake sure that the file isn't open!\n\nExiting...\n")
        exit(1)
    
    
    
def is_empty_row(row):
    """Check if all values in the row are empty, None, or whitespace."""
    for key, value in row.items():
        if isinstance(value, list):
            row[key] = None
    return all((value is None or value.strip() == "") for value in row.values())



# === FUNCTION: SEND EMAIL ===
def send_email(recipient_email, name, subject, body, attachments):
    """
    Sends an email to a single recipient with optional CC and attachments.

    Args:
        recipient_email (str): Primary recipient's email address.
        name (str): Recipient's name for personalization.
        additional_recipients (list): List of CC recipients' email addresses.
        subject (str): Subject of the email.
        body (str): HTML content of the email body.
        attachments (list): List of file paths to be attached.
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
        if len(sys.argv) > 1 and sys.argv[1].strip() == "other_script":
            output_txt_path = os.path.join(NEW_FOLDER_PATH, "output_dir_path.txt")
            
            with open(output_txt_path, "r") as file:
                output_txt_path = file.read()
                
            attachment_path = os.path.join(output_txt_path, attachments)
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
        print("Incorrect Gmail App Password!")
        exit(1)   
    except gaierror as e:
        print("Failed to send Emails....\nCheck your Internet connection\n")
        exit(1)
    except Exception as e:
        # Log failure
        logging.error(f"Failed to send email to {recipient_email}: {e}")
        print(f"Failed to send email to {recipient_email}: {e}")



def check_csv(csv_file_path, attachment_mode, additional_column=None):
        # Open the CSV file
    with open(csv_file_path, "r", encoding="utf-8") as csv_file:
        reader = csv.DictReader(csv_file)
        try:
            # Check if the file is not empty and contains rows after the header
            rows = list(reader)
        except:
            print("\nError in reading CSV file.\nEnsure that the file is not corrupted.\n\nExiting...\n")
            exit(1)

        if not rows or reader.fieldnames is None:  # If rows is empty, only the header or completely empty file
            print("\nThe CSV file is either empty or only contains the header.\n")
            exit(1)
        else:
            # Filter out rows that are effectively empty
            csv_file.seek(0)
            reader = csv.DictReader(csv_file)
            data_rows = []
            
            for row in reader:
                # Check if the row is empty
                if not is_empty_row(row):
                    data_rows.append(row)

            if not data_rows:
                print("\nThe CSV file has no meaningful data rows (only empty rows or whitespace).\n")
                exit(1)

        csv_file.seek(0)
        reader = csv.DictReader(csv_file)
        # Check if required columns exist
        required_columns = {"Full Name", "Email"}
        if additional_column:
            required_columns.add(additional_column)
        
        csv_file.seek(0)
        reader = csv.DictReader(csv_file)
        missing_columns = required_columns - set(reader.fieldnames or [])
        try:
            if missing_columns:
                raise ValueError(f"\nMissing required columns in the CSV file: {', '.join(missing_columns)}\n")

            if attachment_mode in {"Respective", "Common"} and "Attachments" not in reader.fieldnames:
                raise ValueError(f"\nThe 'Attachments' column is required for the selected ATTACHMENT_MODE i.e \'{attachment_mode}\'\n")
        except Exception as e:
            print(f"\nError:{e}")
            exit(1)
        
        csv_file.seek(0)
        reader = csv.DictReader(csv_file)
        empty_rows = []
        # Check each row for empty values
        for row_index, row in enumerate(reader, start=2):
            if not row.get("Full Name", "") is None and not row.get("Email", "") is None:
                if row.get("Full Name", "").strip() == "" or row.get("Email", "").strip() == "":
                    empty_rows.append(row_index)

        if empty_rows:
            print(f"Full Name or Email not found in Row Index - {empty_rows}")
            exit(1)

def check_attachments(csv_file_path, attachment_mode):
    with open(csv_file_path, "r", encoding="utf-8") as csv_file:
        csv_file.seek(0)  # Reset file pointer
        reader = csv.DictReader(csv_file)
        
        is_missing = False
        # Read the common attachments if needed
        if attachment_mode == "Common":
            first_row = next(reader, None)
            if first_row:
                common_attachments = (first_row["Attachments"].split(";"))
                for attachment in common_attachments:
                    if not attachment:
                        is_missing = True
                        print(f"\nThe first row attachment is not specified for the selected Attachment Mode \'Common\'")
                    attachment_path = os.path.join(ATTACHMENTS_DIRECTORY_PATH, attachment)
                    if not os.path.exists(attachment_path):
                        is_missing = True
                        print(f"\nCommon attachment of first row not found - {attachment}")
                        
        elif attachment_mode == "Respective":
            missing_files =[]
            for row_index, row in enumerate(reader, start=2):
                if not row.get("Attachments", "") is None:
                    attachments = row.get("Attachments", "").split(";")
                    missing_files = [path.strip() for path in attachments if path.strip() and not os.path.exists(os.path.join(ATTACHMENTS_DIRECTORY_PATH,path.strip()))]
                if missing_files:
                    is_missing = True                    
                    print(f"Attachment not found - Row Index \'{row_index}\' - {missing_files}")
        
        elif attachment_mode == "Other":
            for row_index, row in enumerate(reader, start=2):
                name = row.get("Full Name", "").title().strip()
                attachments = f"{name.title().strip().replace(' ', '_')}_certificate.pdf"
                
                output_txt_path = os.path.join(NEW_FOLDER_PATH, "output_dir_path.txt")
                with open(output_txt_path, "r") as file:
                    output_txt_path = file.read()
                    
                attachment_path = os.path.join(output_txt_path, attachments)
                if not os.path.exists(attachment_path):
                    is_missing = True                    
                    print(f"Attachment not found: Row Index \'{row_index}\' - {attachment}")
                    
        if is_missing:
            print("\nExiting...\n")
            exit(1)
            


# === FUNCTION: SEND BULK EMAILS ===
def send_bulk_emails(csv_file_path):
    """
    Reads recipient details from a CSV file and sends emails to all recipients.

    Args:
        csv_file_path (str): Path to the CSV file containing recipient details.
    """
    global ATTACHMENT_MODE
    
    try:
        # Read the email body template
        body_template = read_email_body_template()
        if not body_template:
            print("\nEmail body template is empty.\n\nExiting...\n")
            return
            
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
                    if row.get("Email", "") is None or row.get("Full Name", "") is None:
                        raise ValueError("Missing recipient email or name in a row.")
                    else:
                        # Extract recipient details
                        recipient_email = row.get("Email", "").lower().strip()
                        name = row.get("Full Name", "").title().strip()

                    # Determine attachments
                    if ATTACHMENT_MODE == "Respective":
                        attachments = (
                            row.get("Attachments", "").split(";") if row.get("Attachments", "").strip() else []
                        )

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


# === MAIN ENTRY POINT ===
if __name__ == "__main__":
    
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
    
    
    # === CONFIGURATION ===
    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 587
    SENDER_EMAIL = "raqeeb2709@gmail.com"
    EMAIL_SUBJECT = "Subject"

    # === ATTACHMENT MODE ===
    # 'None': No attachments will be sent.
    # 'Common': The first recipient's attachments will be sent to everyone.
    # 'Respective': Attachments from the CSV file will be used for each recipient respectively.
    ATTACHMENT_MODE = "None"  # Change this to 'None', 'Common', or 'Respective' (in quotes)
    
    ATTACHMENTS_DIRECTORY_PATH = os.path.join(EMAIL_SENDER_DIRECTORY_PATH, "Attachments")
    SPREADSHEET_DIRECTORY_PATH = os.path.join(EMAIL_SENDER_DIRECTORY_PATH, "Spreadsheet")
    os.makedirs(SPREADSHEET_DIRECTORY_PATH, exist_ok=True)
    
    csv_files = get_files(SPREADSHEET_DIRECTORY_PATH, 'CSV')
    spreadsheet_file = get_single_file('Spreadsheet', SPREADSHEET_DIRECTORY_PATH, 'CSV')
    CSV_FILE_PATH = os.path.join(SPREADSHEET_DIRECTORY_PATH, spreadsheet_file)
    
    if len(sys.argv) > 1 and sys.argv[1] == "other_script":
        NEW_FOLDER_PATH = os.path.join(ROOT_REPO_PATH, "New_Folder")
        DIR_PATH = NEW_FOLDER_PATH
        CSV_FILE_PATH = os.path.join(NEW_FOLDER_PATH, "tosend.csv")
        ATTACHMENT_MODE = "Other"
        EMAIL_SUBJECT == sys.argv[2]
    else:
        DIR_PATH = EMAIL_SENDER_DIRECTORY_PATH

    BODY_TEMPLATE_FILE_PATH = os.path.join(DIR_PATH, "body_template.html")
    LOG_FILE_PATH = os.path.join(DIR_PATH, "email_log.txt")
    GMAIL_APP_PASSWORD_FILE_PATH  = os.path.join(DIR_PATH, "gmail_app_password.txt")
    
    os.makedirs(ATTACHMENTS_DIRECTORY_PATH, exist_ok=True)
    for file in [BODY_TEMPLATE_FILE_PATH, LOG_FILE_PATH, GMAIL_APP_PASSWORD_FILE_PATH]:
        if not os.path.exists(file):
            with open(file, 'w') as f:
                pass

    # === SET UP LOGGING ===
    logging.basicConfig(
        filename=LOG_FILE_PATH,
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
    )
    
    if not os.path.exists(CSV_FILE_PATH):
        print("\n\'recipients.csv\' file not found!\nCreating and Initializing the file with header row.\n")
        with open(CSV_FILE_PATH, mode='w', newline='') as csv_file:
            csv_writer = csv.writer(csv_file)
            csv_writer.writerow(['Full Name', 'Email', 'Attachments'])  # Write header row
        print("\nExiting....\n")
        exit(1)
    
    try:
        # Read the names from the file
        with open(GMAIL_APP_PASSWORD_FILE_PATH, 'r') as file:
            lines = file.readlines()
            password = [line.strip() for line in lines if line.strip()]
            if len(password) != 1 or password[0].count(" ") != 0:
                print("\nError in gmail app password!!\nInput password in a single line and without any spaces.\n\nExiting...\n")
                exit(1)
        SENDER_PASSWORD = password[0]
            
    except FileNotFoundError:
        print("\nError in reading password file!\nPlease ensure that the file exists at correct location.\n\nExiting...\n")
        exit(1)
    except Exception as e:
        print(f"\nError in reading password file!\n{e}\nPlease ensure that the file is not corrupted.\n\nExiting...\n")
        exit(1)

    sort_csv(CSV_FILE_PATH)
    # Check command-line arguments
    if not (len(sys.argv) > 1 and sys.argv[1] == "other_script"):
        check_csv(CSV_FILE_PATH, ATTACHMENT_MODE)
    
    check_attachments(CSV_FILE_PATH, ATTACHMENT_MODE)
                
    send_bulk_emails(CSV_FILE_PATH)
