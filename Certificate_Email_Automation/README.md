# Certificate Email Automation Script

This script automates the process of extracting recipient details, generating certificates, and sending personalized emails with certificates as attachments.

## Features

- **Spreadsheet Extraction**: Reads a CSV spreadsheet and filters attendees based on their "Attendance" status.
- **Certificate Generation**: Integrates with a certificate generation script to create personalized certificates.
- **Email Sending**: Sends certificates via email with a customizable HTML email template.
- **File Validation**: Ensures required directories, files, and inputs are properly structured.
- **Error Logging**: Logs all actions and errors in a log file for easy debugging.

---

## Prerequisites

- **Python**: Ensure Python 3.6 or above is installed.
- **Required Modules**: The script relies on utility functions from `Utilities.utils`.
- **Gmail App Password**: Generate a Gmail App Password for secure email authentication. (Refer to Gmail documentation for steps to create an app password [here](https://knowledge.workspace.google.com/kb/how-to-create-app-passwords-000009237)).

---

## Repository Structure

```plaintext
Resources/
│
├── Certificate_Email_Automation/
│   │
│   ├── Certificate_Template/
│   │
│   ├── Spreadsheet/
│   │
│   ├── Wordlist/
│   │
│   ├── Generated_Certificates/ (created automatically for certificates)
│   │
│   ├── body_template.html
│   │
│   ├── email_log.txt
│   │
│   └── gmail_app_password.txt
│
└── extract_certify_and_email.py
```

---

## Usage

### 1. Prepare the Environment
> Run the script once to create all the necessary files and directories using the below command:

   ```bash
   python extract_certify_and_email.py
   ```

1. Place your spreadsheet file in the `Spreadsheet` directory. Ensure it has the following columns:
    > Ensure that the column names doesn't have any preceeding or succeeding whitespaces.
   - **Full Name**: Full name of the recipient.
   - **Email**: Email address of the recipient.
   - **Attendance**: A column to indicate attendance (use `TRUE` for attendees).

   ```csv
   Full Name,Email,Attendance
   ```

2. Place your email body template (`body_template.html`) in the Certificate_Email_Automation directory.
3. Ensure a `.pdf` certificate template is present in the Certificate_Template directory.
4. Add your Gmail App Password to `gmail_app_password.txt` for secure email authentication.
5. Open the terminal and navigate to the project directory and run the script:

  ```bash
  python extract_certify_and_email.py
  ```

### 2. Workflow Overview

- **Spreadsheet Extraction**: The script extracts the "Full Name" and "Email" columns for attendees marked as `TRUE` in the "Attendance" column.
  - Creates a `tosend.csv` file with these details.
  - Writes the "Full Name" column to `wordlist.txt` for certificate generation.

- **Certificate Generation**: Calls `certificate_generator.py` to generate personalized certificates.

- **Email Sending**: Calls `send_email.py` to send emails with the generated certificates attached.

---

## Error Handling and Logging

- Logs are saved in the `email_log.txt` file.
- Common issues include:
  - **Authentication Error**: Check your Gmail App Password.
  - **File Not Found**: Ensure all required files and directories exist.
  - **KeyError**: Ensure the input spreadsheet contains the required columns ("Full Name", "Email", "Attendance").

---

## Customization

- **Email Subject**: Update the `EMAIL_SUBJECT` variable in the script.
- **Gmail Credentials**: Replace `SENDER_EMAIL` with your Gmail address in the `send_email.py` script.

---

## Troubleshooting

- **Invalid CSV Format**: Ensure the input spreadsheet has the required columns.
- **Template Issues**: Verify that the email body template (`body_template.html`) is correctly formatted.

---

## Dependencies

- **Utilities.utils** module for file handling and validation.
- **External Scripts**:
  - `certificate_generator.py` for generating certificates.
  - `send_email.py` for sending emails.
