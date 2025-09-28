# Google Form Email Automation with Error Logging

This project automates email responses to Google Form submissions. Using Apps Script, it validates email addresses, sends personalized emails with an HTML template, and logs errors for debugging.

---

## Features

- **Personalized HTML Emails**: Sends customized emails using a predefined HTML template.
- **Email Validation**: Ensures that only valid email addresses receive responses.
- **Error Logging**: Logs errors to a dedicated sheet for easy debugging.
- **Trigger Setup**: Automatically links the script to form submissions.

---

## Step 1: Create the Google Form

### Design the Form
1. Open [Google Forms](https://forms.google.com) and select **Blank Form** to create a new form.  
![Create Blank Form](assets/images/create-blank-form.png "Create Blank Form")

2. Design the form as per your requirements. Ensure that you add the **two required fields** with the exact names as below and mark them as **Required**:
   - **Full Name**
   - **Email**  
  ![required questions](assets/images/required-questions.png "required questions")

3. Ensure that in **Form Settings**, the option to "Collect email addresses" is set to **Don’t collect**  
  > This setting avoids duplicating email fields and ensures compatibility with the script.  

  _In Settings tab:_ &nbsp;**Settings > Responses > Collect email addresses**  
  ![response settings](assets/images/response-settings(1).png "response settings")  

  _In Settings tab:_ &nbsp;**Defaults > Form defaults > Collect email addresses by default**  
  ![response default settings](assets/images/response-settings(2).png "default settings")

---

## Step 2: Link the Form to a Google Sheet

1. In the Google Form, go to the **Responses** tab.
2. Click the green spreadsheet icon to create a linked Google Sheet where responses will be stored.  
![Link Spreadsheet](assets/images/link-spreadsheet.png "Link Spreadsheet")
3. Choose to create a new spreadsheet and provide a name for it, then click **Create**. (You can also select an existing spreadsheet if preferred.)  
![Response destination selection](assets/images/response-dest-selection.png "Response destination selection")

---

## Step 3: Add the Email Script to Google Sheets

### Open the Script Editor
1. Open the linked Google Sheet and navigate to **Extensions > Apps Script**.  
![Apps Script navigation](assets/images/appscript-navigation.png "Apps Script navigation")

### Import the Script
1. Open the `sheet_appscript.js` file in this repository to view the script. ( [Direct Link to script](./sheet_appscript.js) )
2. Copy the contents of the file.
3. Clear any existing code in the Apps Script editor and paste the copied script into the editor.  
![Paste script](assets/images/paste-script.png "Paste script")

---

## Step 4: Configure the Script

### Change the Email Subject and HTML Body content
1. Edit the `SUBJECT` constant in the script to your desired email subject.
2. Modify the `getEmailTemplate()` function to customize the email body content as needed.  
![Apps Script customization](assets/images/appscript-customization.png "Apps Script customization")

### Rename and Save
1. Click on the project name (default is "Untitled project") in the top left corner and rename it (e.g., `Email Automation Script`).  
![Rename Apps Script](assets/images/rename-appscript.png "Rename Apps Script")
2. Save the script (`Ctrl+S`).

---

## Step 5: Setup Trigger and Authorize the Script

### Set Up the Trigger
1. In the Apps Script editor toolbar, click on the dropdown `getEmailTemplate` and select `setupTrigger`.
2. Click the **Run (▶️)** icon to run the function. This sets up a trigger to automate email sending.  
![Trigger function](assets/images/trigger-function.png "Trigger function")

### Authorize the Script
1. A dialog will appear requesting authorization. Click **Review Permissions**.  
![Authorize](assets/images/authorize.png "Authorize permissions")  
2. Choose your Google account.
3. If you see a warning about the app being unverified, click on **Advanced** and then **Go to {Project Name} (unsafe)**.  
![Advanced options](assets/images/advanced-options.png "Advanced options")
4. Choose **Select all** permissions and click **Continue**.  
![Grant access](assets/images/grant-access.png "Grant access")

---

## Step 6: Copy the Shortened URL of the Form
1. Go back to the Google Form.
2. Click the **Publish** button at the top right.  
![Publish form](assets/images/publish-form.png "Publish form")
3. Set the Responders to **Anyone with the link** option and click **Publish**.  
![Manage responders](assets/images/manage-responders.png "Manage responders")
4. Check the **Shorten URL** checkbox and **Copy** the link provided.  
![Copy shortened url](assets/images/copy-shortened-url.png "Copy shortened url")
5. Share this link with your audience to start collecting responses.

---

## Step 7: Test the Automation

1. Submit a test response using the Google Form.
2. Check the test email inbox for the personalized response.
3. If issues occur, refer to the **Error Log** at the bottom of the Google Sheet for error logs for troubleshooting.  
![View error log](assets/images/view-error-log.png "View error log")

---

## Notes

- **Email Template**: Customize the email subject by modifying the `SUBJECT` constant and body in the `getEmailTemplate()` function within `appscript.js`.
- **Email Validation**: The `isValidEmail()` function checks for proper email formatting.
- **Error Logging**: Errors are logged to a sheet named **"Error Log"** for debugging.
- **Email Index**: Ensure `e.values[1]` corresponds to the email column in your sheet.
- **Gmail Limits**: Free Gmail accounts are limited to **100 emails per day**.
