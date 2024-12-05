import csv
import logging
import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

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

        # Add attachments
        for attachment_path in attachments:
            if attachment_path.strip():  # Ensure path is not empty
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

        # Connect to Gmail's SMTP server and send the email
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(
                SENDER_EMAIL,
                [recipient_email],
                msg.as_string(),
            )

        # Log success
        logging.info(f"Email sent to {recipient_email}")
        print(f"Email sent to {recipient_email}")

    except Exception as e:
        # Log failure
        logging.error(f"Failed to send email to {recipient_email}: {e}")
        print(f"Failed to send email to {recipient_email}: {e}")


# === FUNCTION: SEND BULK EMAILS ===
def send_bulk_emails(csv_file_path):
    """
    Reads recipient details from a CSV file and sends emails to all recipients.

    Args:
        csv_file_path (str): Path to the CSV file containing recipient details.
    """
    try:
        # Read the email body template
        body_template = read_email_body_template()
        if not body_template:
            print("Email body template is empty. Exiting...")
            return

        # Open the CSV file
        with open(csv_file_path, "r", encoding="utf-8") as csv_file:
            reader = csv.DictReader(csv_file)

            # Determine common attachments if needed
            common_attachments = []
            if ATTACHMENT_MODE == "Common":
                first_row = next(reader, None)
                if first_row:
                    common_attachments = (
                        first_row["Attachments"].split(";")
                        if first_row.get("Attachments")
                        else []
                    )

        # Open the CSV file
        with open(csv_file_path, "r", encoding="utf-8") as csv_file:
            reader = csv.DictReader(csv_file)
            
            # Process each recipient
            for row in reader:
                # Extract recipient details
                recipient_email = row["Email"]
                name = row["Name"]

                if ATTACHMENT_MODE == "Respective":
                    attachments = (
                        row["Attachments"].split(";") if row.get("Attachments") else []
                    )
                elif ATTACHMENT_MODE == "Common":
                    attachments = common_attachments
                else:  # ATTACHMENT_MODE == "None"
                    attachments = []

                # Customize the email body
                personalized_body = body_template.replace("{{name}}", name)

                # Send the email
                send_email(
                    recipient_email, name, EMAIL_SUBJECT, personalized_body, attachments
                )

    except FileNotFoundError:
        logging.error(f"CSV file not found: {csv_file_path}")
        print(f"CSV file not found: {csv_file_path}")
    except Exception as e:
        logging.error(f"Error while processing CSV file: {e}")
        print(f"Error while processing CSV file: {e}")


# === MAIN ENTRY POINT ===
if __name__ == "__main__":
    
    print("\n" + " Email Sender ".center(24, "-"))
    
    if os.getcwd()[-9:] == "Resources":
        # EMAIL_SENDER_DIRECTORY_PATH = os.path.join(os.getcwd(), 'Email_Sender')
        print("\nTo run this script, please change your working directory to the \'Email_Sender\' directory using the command \'cd Email_Sender\'.\n\nExiting...\n")
        exit(1)

    elif os.getcwd()[-12:] == "Email_Sender":
        EMAIL_SENDER_DIRECTORY_PATH = os.getcwd()
        ROOT_REPO_PATH = os.path.join(os.getcwd(), '..')
        
    else:
        print("\nPlease change your working directory to the main repository.\n\nExiting...\n")
        exit(1)
    
    
    # === CONFIGURATION ===
    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 587
    SENDER_EMAIL = "cyberelites.nsakcet@gmail.com"
    EMAIL_SUBJECT = "Subject"

    BODY_TEMPLATE_FILE_PATH = os.path.join(EMAIL_SENDER_DIRECTORY_PATH, "body_template.html")
    LOG_FILE_PATH = os.path.join(EMAIL_SENDER_DIRECTORY_PATH, "email_log.txt")
    CSV_FILE_PATH = os.path.join(EMAIL_SENDER_DIRECTORY_PATH, "recipients.csv")
    GMAIL_APP_PASSWORD_FILE_PATH  = os.path.join(EMAIL_SENDER_DIRECTORY_PATH, "gmail_app_password.txt")

    # === ATTACHMENT MODE ===
    # 'None': No attachments will be sent.
    # 'Common': The first recipient's attachments will be sent to everyone.
    # 'Respective': Attachments from the CSV file will be used for each recipient respectively.
    ATTACHMENT_MODE = "Respective"  # Change this to 'None', 'Common', or 'Respective' (in quotes)

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
            csv_writer.writerow(['Name', 'Email', 'Attachments'])  # Write header row
        print("\nExiting....\n")
        exit(1)
    
    try:
        # Read the names from the file
        with open(GMAIL_APP_PASSWORD_FILE_PATH, 'r') as file:
            lines = file.readlines()
            password = [line.strip() for line in lines if line.strip()]
            if len(password) != 1 or password[0].count(" ") != 0:
                print("Error in gmail app password!!!")
                exit(1)
        SENDER_PASSWORD = password[0]
            
    except FileNotFoundError:
        print("\nError in reading password file!\nPlease ensure that the file exists at correct location.\n\nExiting...\n")
        exit(1)
    except:
        print("\nError in reading password file!\nPlease ensure that the file is not corrupted.\n\nExiting...\n")
        exit(1)

    send_bulk_emails(CSV_FILE_PATH)
