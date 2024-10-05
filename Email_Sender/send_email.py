import logging
import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import List, Optional

# logging setup for debug and error tracking
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class EmailSender:
    """
    A class to send emails via Gmail's SMTP server using App Passwords for authentication.
    """
    def __init__(self, smtp_server: str, smtp_port: int, gmail_user: str, gmail_password: str):
        """
        Initializes the EmailSender with SMTP settings and credentials.
        :param smtp_server: SMTP server address (Gmail).
        :param smtp_port: Port for SMTP server.
        :param gmail_user: Gmail username (email address).
        :param gmail_password: Gmail App Password (or account password if 2FA is disabled).
        """
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.gmail_user = gmail_user
        self.gmail_password = gmail_password

    def send_email(self, subject: str, body: str, to_emails: List[str], 
                   attachments: Optional[List[str]] = None):
        """
        Sends an email with the specified subject, body, and optional attachments.
        :param subject: Subject of the email.
        :param body: Body of the email (plain text or HTML).
        :param to_emails: List of recipient email addresses.
        :param attachments: List of file paths to attach to the email (optional).
        """
        try:
            # Create a MIME multipart message
            message = MIMEMultipart()
            message['From'] = self.gmail_user
            message['To'] = ', '.join(to_emails)
            message['Subject'] = subject

            # Attach the body of the email
            message.attach(MIMEText(body, 'plain'))

            # Attach any files if provided
            if attachments:
                for attachment in attachments:
                    self._attach_file(message, attachment)

            # Connect to the Gmail SMTP server and send the email
            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()  # Upgrade the connection to secure encrypted TLS
                server.login(self.gmail_user, self.gmail_password)
                text = message.as_string()
                server.sendmail(self.gmail_user, to_emails, text)

            logger.info("Email sent successfully.")
        
        except Exception as e:
            logger.error(f"Failed to send email: {str(e)}")

    def _attach_file(self, message: MIMEMultipart, file_path: str):
        """
        Attach a file to the MIME email message.
        :param message: The MIME message object.
        :param file_path: Path to the file to attach.
        """
        try:
            with open(file_path, "rb") as file:
                mime_base = MIMEBase('application', 'octet-stream')
                mime_base.set_payload(file.read())
                encoders.encode_base64(mime_base)
                mime_base.add_header('Content-Disposition', f'attachment; filename={os.path.basename(file_path)}')
                message.attach(mime_base)

            logger.info(f"Attached file: {file_path}")
        except Exception as e:
            logger.error(f"Failed to attach file {file_path}: {str(e)}")


def main():
    """
    Main function to set up email parameters and send an email.
    """
    # SMTP settings for Gmail
    SMTP_SERVER = "smtp.gmail.com"
    SMTP_PORT = 587  # Port for TLS

    # Email account credentials
    GMAIL_USER = ""
    GMAIL_PASSWORD = ""  # Gmail App Password for security

    # Email content
    SUBJECT = ""
    BODY = ""
    TO_EMAILS = []  # List of recipients
    ATTACHMENTS = []  # Optional list of attachments (.txt,.pdf and others)

    # Initialize the email sender
    email_sender = EmailSender(SMTP_SERVER, SMTP_PORT, GMAIL_USER, GMAIL_PASSWORD)

    # Send the email
    email_sender.send_email(subject=SUBJECT, body=BODY, to_emails=TO_EMAILS, attachments=ATTACHMENTS)


if __name__ == "__main__":
    main()