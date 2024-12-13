# Email Automation Script

This script is designed to send bulk emails with optional attachments using Gmail's SMTP server. It provides features such as email personalization, attachment handling, and integration with certificate automation workflows.

---

## Features

- **Bulk Email Sending**: Reads recipient details from a CSV file and sends personalized emails.
- **Customizable Email Body**: Uses an HTML template for email body content.
- **Attachment Modes**: Supports the following attachment modes:
  - **None**: No attachments.
  - **Common**: Same attachments for all recipients.
  - **Respective**: Attachments specified individually for each recipient in the CSV file.
  - **Other**: Specific attachment naming convention for certificate automation workflows.
- **Automation Integration**: Integrates with certificate generation scripts to automatically attach and send certificates.
- **Error Logging**: Logs all operations, including errors, for debugging.

---

## Prerequisites

1. **Python**: Ensure `Python 3.6` or above is installed.
2. **Gmail App Password**: Generate a Gmail App Password for secure authentication. (Refer to Gmail documentation for steps to create an app password [here](https://knowledge.workspace.google.com/kb/how-to-create-app-passwords-000009237)).
3. **Files and Directories**:
   - **CSV File**: A CSV file with recipient details.
   - **HTML Template**: A body template (`body_template.html`) for email content.
   - **Attachments**: Directory for storing attachments.

---

## Setup
> Run the script once to create all the necessary files and directories from the Email_Sender directory using the below command:

   ```bash
   python send_email.py
   ```

### 1. Repository Structure
The repository should follow this structure:

```plaintext
Email_Sender/
│
├── Attachments/
│
├── Spreadsheet/
│
├── body_template.html
│
├── gmail_app_password.txt
│
├── email_log.txt
│
└── email_sender.py
```

---

### 2. Required Files

#### CSV File
Should include the following columns:
> Ensure that the column names doesn't have any preceeding or succeeding whitespaces.
- **Email**: Recipient's email address.
- **Full Name**: Recipient's full name for personalization.
- **Attachments (optional)**: Semicolon-separated list (`;`)of attachment filenames.

```csv
Full Name,Email,Attachments
```

#### HTML Body Template (`body_template.html`)
A template file with placeholders like {{name}} for personalization.

```html
<html>
    <body>
        <!-- Body Contents -->
    </body>
</html>
```

#### Gmail App Password File (`gmail_app_password.txt`)
Contains your Gmail App Password.

---

## Usage

1. Run the script once to create all the necessary files and directories.
2. Place your **CSV file** in the `Spreadsheet` directory.
3. Place your **HTML template** within the `Email_Sender` directory.
4. Place any **attachments** (if needed) in the `Attachments` directory.
5. Open the terminal and navigate to the `Email_Sender` directory.
6. Choose the appropriate attachment mode in the script:
    - 'None': No attachments.
    - 'Common': Attachments from the first row will be sent to all recipients.
    - 'Respective': Individual attachments from the Attachments column.
7. Run the script:

   ```bash
   python send_email.py
   ```

---

## Error Handling and Logging

- Logs are saved in the `email_log.txt` file.
- Common issues include:
  - **Authentication Error**: Check your Gmail App Password.
  - **File Not Found**: Ensure all specified files and directories exist.
  - **Network Issues**: Verify your internet connection.

---

## Customization

- **Email Subject**: Update the `EMAIL_SUBJECT` variable in the script.

---

## Troubleshooting

- **Invalid CSV Format**: Ensure the CSV has the correct columns.
- **Template Issues**: Verify that placeholders like `{{name}}` match your CSV column names.
