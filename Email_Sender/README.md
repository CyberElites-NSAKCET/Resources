# Email Automation Script

This script sends bulk personalized emails using Gmail's SMTP server. It supports HTML templates, multiple attachment modes, and integrates with certificate automation workflows.

---

## Features

- **Bulk Email Sending**: Reads recipient details from a CSV file and sends personalized emails.
- **Customizable Email Body**: Uses an HTML template (`body_template.html`). You may use `{{name}}` as a placeholder for personalization, or omit it if not needed.
- **Attachment Modes**:
  - **None**: No attachments.
  - **Common**: Same attachments for all recipients (from the first row).
  - **Respective**: Attachments specified per recipient in the CSV.
  - **Other**: Used for certificate automation; attaches generated certificates by name.
- **Automation Integration**: Can be called by automation scripts for certificate distribution.
- **Error Logging**: Logs all operations and errors to `email_log.txt`.

---

## Prerequisites

1. **Python 3.9+**
2. **Gmail App Password**: [How to create one](https://knowledge.workspace.google.com/kb/how-to-create-app-passwords-000009237)
3. **Required Python Packages**:  
   - No extra packages beyond the Python standard library are required for email sending.  
   - For certificate automation, see the relevant generator's requirements.

---

## Configuration

⚠️ **Important:** Edit `config.json` in the project root to set:

- `smtp_server`: e.g., `smtp.gmail.com`
- `smtp_port`: e.g., `587`
- `sender_email`: Your Gmail address
- `gmail_app_password`: Your Gmail App Password (16 characters, no spaces)
- `email_subject`: Subject line for emails
- `attachment_mode`: `"None"`, `"Common"`, `"Respective"`, or `"Other"`

---

## Setup

> **💡Run the script once to create all necessary files and directories:**

1) Clone the repo
   ```bash
   git clone https://github.com/CyberElites-NSAKCET/Resources 
   ```  
2) Change the working directory to the Email Sender directory
   ```bash
   cd Resources
   cd Email_Sender
   ```  
3) Run the script once to create all the files and directories 
    ```bash
    python send_email.py
    ```

### Directory Structure

```
Email_Sender/
│
├── Attachments/
│   └── <attachment files>
│
├── Spreadsheet/
│   └── <spreadsheet_file>.csv
│
├── body_template.html
│
├── email_log.txt
│
└── send_email.py
```

---

## Required Files

### 1) CSV File

- Place your CSV file in the `Spreadsheet/` directory.
- Required columns:  
  - `Full Name`
  - `Email`
  - `Attachments` (optional; semicolon-separated for multiple files)

Example:
```csv
Full Name,Email,Attachments
Abdul,abd@example.com,file1.pdf;file2.pdf
Raqeeb,raq@example.com,
```

### 2) HTML Body Template (`body_template.html`)

- Place in the `Email_Sender/` directory.
- You may use `{{name}}` as a placeholder for personalization, or omit it if not needed.  
- The template can be any valid HTML content as per your requirements.

Example with `{{name}}`:
```html
<html>
  <body>
    <p>Hello {{name}},</p>
    <p>Welcome to our event!</p>
  </body>
</html>
```

Example without `{{name}}`:
```html
<html>
  <body>
    <p>Thank you for being part of our community!</p>
    <p>Best regards,<br>CyberElites Club</p>
  </body>
</html>
```

### 3) Attachment files *(optional)*
 - Place them in the `Attachments/` directory.
 - Ensure that you have specified the right Attachment mode in the `config.json` file as per your requirements.

---

## Usage

1. Place your CSV in `Spreadsheet/`, HTML template in `Email_Sender/`, and any attachments in `Attachments/`.
2. Set the desired attachment mode in `config.json`.
3. Run the script:

    ```bash
    python send_email.py
    ```

- For automation (certificate workflow), the script is called with an argument by the automation script and uses the `"Other"` mode.

---

## Error Handling and Logging

- Logs are saved in `email_log.txt`.
- Common issues:
  - **Authentication Error**: Check your Gmail App Password.
  - **File Not Found**: Ensure all files and directories exist.
  - **Invalid CSV Format**: Ensure required columns are present.
  - **Network Issues**: Check your internet connection.

---

## Customization

- **Email Subject**: Set in `config.json` as `email_subject`.
- **HTML Template**: Edit `body_template.html` as desired. Use `{{name}}` for personalization if needed, or omit it.

---

## Troubleshooting

- **Invalid CSV Format**: Ensure the CSV has the correct columns and no extra whitespace in headers.
- **Template Issues**: If using `{{name}}`, ensure it matches the CSV column. Otherwise, you may omit it.
- **Attachment Issues**: Ensure files listed in `Attachments` exist in the `Attachments/` directory.

---
