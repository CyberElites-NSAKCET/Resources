# Certificate Email Automation Script

This script automates the process of extracting recipient details from a spreadsheet, generating personalized certificates, and sending personalized emails with customizable HTML templates and certificates as attachments.

## Features

- **Spreadsheet Extraction**: Reads a CSV spreadsheet and filters attendees based on their `"Attendance"` status.
- **Certificate Generation**: Integrates with a `certificate_generator.py` script to create personalized certificates.
- **Email Sending**: Sends certificates via email with a customizable HTML email template.
- **File Validation**: Ensures required directories, files, and inputs are properly structured.
- **Configurable Workflow**: Uses `config.json` for Gmail credentials, subject lines, and integration settings.
- **Error Logging**: Logs all actions and errors in a `email_log.txt` file for easy debugging.

---

## Prerequisites

- **Python**: Ensure Python 3.6 or above is installed.
- **Gmail App Password**: [How to create one](https://knowledge.workspace.google.com/kb/how-to-create-app-passwords-000009237).

---

## Configuration
> ⚠️ Important: Edit config.json in the project root to set:

 - `smtp_server`: e.g., `smtp.gmail.com`
 - `smtp_port`: e.g., `587`
 - `sender_email`: Your Gmail address
 - `gmail_app_password`: Your Gmail App Password (16 characters, no spaces)
 - `email_subject`: Subject line for certificate emails
 - `attachment_mode`: Always `"Other"` (certificate automation mode)

---

## Setup
> **💡Run the script once to create all necessary files and directories:**  

1) Clone the repo  
    ```bash
    git clone https://github.com/CyberElites-NSAKCET/Resources
    cd Resources
    ```

2) Install the dependencies using the command:  
    ```bash
    pip install -r requirements.txt
    ```

3) Run the script once to create all the files and directories  
    ```bash
    python extract_certify_and_email.py
    ```

## Repository Structure

```plaintext
Resources/
│
├── extract_certify_and_email.py
│
└── Certificate_Email_Automation/
    │
    ├── Certificate_Template/
    │   └── <certificate_template>.pdf
    │
    ├── Spreadsheet/
    │   └── <spreadsheet_file>.csv
    │
    ├── Wordlist/
    │
    ├── Generated_Certificates/ (created automatically for certificates)
    │
    ├── body_template.html
    │
    └── email_log.txt
```

---

## Required Files

### 1) Spreadsheet (CSV)

- Place your CSV file in the `Spreadsheet/` directory.
- Required columns:
  > ⚠️ Ensure that the column names doesn't have any preceeding or succeeding whitespaces.
   - `Full Name`: Full name of the recipient.
   - `Email`: Email address of the recipient.
   - `Attendance`: A column to indicate attendance (use `TRUE` for attendees).

    Example:
    ```csv
    Full Name,Email,Attendance
    Abdul,abd@example.com,TRUE
    Raqeeb,raq@example.com,FALSE
    ```

### 2) HTML Email Template (`body_template.html`)

- Place in `Certificate_Email_Automation/`.
- You may use `{{name}}` as a placeholder for personalization, or omit it if not needed.  
- The template can be any valid HTML content as per your requirements.

    Example with `{{name}}`:
    ```html
    <html>
      <body>
        <p>Dear {{name}},</p>
        <p>Thank you for attending our event. Your certificate is attached.</p>
      </body>
    </html>
    ```
    Example without `{{name}}`:
    ```html
    <html>
      <body>
        <p>Dear Attendee,</p>
        <p>Thank you for attending our event. Your certificate is attached.</p>
      </body>
    </html>
    ```

### 3) Certificate Template

- Place your `.pdf` certificate template in the `Certificate_Template/` directory.
- Names from `wordlist.txt` will be used to generate personalized certificates.

---

## Usage

1) Place your spreadsheet in `Spreadsheet/`, certificate template in `Certificate_Template/`, and ensure `body_template.html` exists.

2) Update `config.json` with Gmail credentials and subject.

3) Run the script:

    ```bash
    python extract_certify_and_email.py
    ```

- The script will:
    - Extract recipients marked `"TRUE"` in Attendance.
    - Generate certificates.
    - Send personalized emails with certificates attached.

### Workflow Overview

- **Spreadsheet Extraction**: The script extracts the `"Full Name"` and `"Email"` columns for attendees marked as `TRUE` in the `"Attendance"` column.
  - Creates a `tosend.csv` file with these details.
  - Writes the `"Full Name"` column to `wordlist.txt` for certificate generation.

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

- **Email Subject**: Set the `email_subject` in the `config.json` file.
- **Certificate Design**: Provide the template PDF in `Certificate_Template/` directory.
- **Email Body**: Edit `body_template.html` file as needed.

---

## Troubleshooting

- **Invalid CSV Format**: Ensure the input spreadsheet has the required columns.
- **Template Issues**: Verify that the email body template (`body_template.html`) is correctly formatted.
- **Email Not Sent**: Verify SMTP settings and internet connection.

---

## Dependencies

- **Utilities.utils** module for file handling and validation.
- **External Scripts**:
  - `certificate_generator.py` for generating certificates.
  - `send_email.py` for sending emails.
