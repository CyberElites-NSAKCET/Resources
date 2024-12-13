// Define the subject line for the email to be sent.
const SUBJECT = "Thank You for Your Submission!";

/**
 * Generate an HTML email template with personalized content.
 * Pass the extracted values (e.g., email, name) to this function.
 * @param {string} email - The email address of the recipient.
 * @param {string} fullname - The full name of the recipient.
 * @returns {string} - The HTML email content.
 */
function getEmailTemplate(fullname) {
    return `
        <html>
            <body>
                <p><strong>Hello ${fullname}</strong>,</p>
                <p>Here's the link to the club website <em><a href="https://cyberelites.org" target="_blank">CyberElites</a></em>. Explore more about us here.</p>
                <p>Thank you!</p>
                <p>Best regards,<br><strong>CyberElites Club</strong></p>
            </body>
        </html>
    `;
}

/**
 * Validate the format of an email address.
 * @param {string} email - The email address to validate.
 * @returns {boolean} - True if the email is valid, otherwise false.
 */
function isValidEmail(email) {
    var regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/; // Standard regex for email validation.
    return regex.test(email);
}

/**
 * Log errors into a spreadsheet for troubleshooting.
 * If the log sheet does not exist, it creates one.
 * @param {Error} error - The error object containing the error message.
 * @param {Date} timestamp - The timestamp of when the error occurred.
 * @param {Array} formData - The submitted form data that caused the error.
 */
function logError(error, timestamp, formData) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("Error Log Sheet");

    // Create a new sheet for logging if it doesn't exist
    if (!logSheet) {
        logSheet = ss.insertSheet("Error Log");
        logSheet.appendRow(["Timestamp", "Error Message", "Form Data"]); // Header row for log entries
    }

    // Append the error details to the log sheet
    logSheet.appendRow([timestamp, error.message, JSON.stringify(formData)]);
}

/**
 * Extract values from the form submission using column names.
 * Pass these extracted values to the getEmailTemplate function.
 * Validate the email format and send the email if valid.
 * If an error occurs, log it in the "Error Log Sheet".
 * @param {Object} e - The event object containing form submission data.
 */
function sendEmailOnSubmit(e) {
    try {
        // Extract values as per the columns in the sheet
        //var timestamp = e.values[0];
        var email = e.values[1];    // *required for sending email to the recipient
        var fullname = e.values[2];
        //var phone = e.values[3];

        // Validate the email format
        if (!isValidEmail(email)) {
            throw new Error("Invalid email format."); // Throw an error if the email is invalid.
        }

        var subject = SUBJECT;

        // Pass extracted values to the email template funtion
        var body = getEmailTemplate(fullname);

        // Send the email using GmailApp
        GmailApp.sendEmail(email, subject, "", { htmlBody: body });

    } catch (error) {
        // Log any errors to the "Error Log Sheet"
        logError(error, new Date(), e.values);
    }
}

/**
 * Setup a trigger for the onFormSubmit event.
 * Ensures that the trigger exists and avoids creating duplicate triggers.
 */
function setupTrigger() {
    // Get all existing triggers for the project
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
        // Check if the trigger for the function "sendEmailOnSubmit" already exists
        if (triggers[i].getHandlerFunction() === "sendEmailOnSubmit") {
            console.log("Trigger already exists.");
            return; // Exit if the trigger is already set.
        }
    }

    // Create a new trigger for the form submission
    ScriptApp.newTrigger("sendEmailOnSubmit")
        .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
        .onFormSubmit()
        .create();
}
