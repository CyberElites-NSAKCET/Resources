const SUBJECT = "Thank You for Your Submission!";

function getEmailTemplate(email) {
    return `
        <html>
            <body>
                <p><strong>Hello Name</strong>,</p>
                <p>Your Email recorded as: ${email}</p>
                <p>Your Name recorded as: Name</p>
                <p>Here's the link to club website <em><a href="https://cyberelites.org" target="_blank">CyberElites</a></em>. Explore more about us here.</p>
                <p>Thank you!</p>
                <p>Best regards,<br><strong>CyberElites Club</strong></p>
            </body>
        </html>
    `;
}

function isValidEmail(email) {
    var regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return regex.test(email);
}

function logError(error, timestamp, formData) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("Error Log Sheet");

    // Create the sheet if it doesn't exist
    if (!logSheet) {
        logSheet = ss.insertSheet("Error Log");
        logSheet.appendRow(["Timestamp", "Error Message", "Form Data"]);
    }

    // Log the error
    logSheet.appendRow([timestamp, error.message, JSON.stringify(formData)]);
}

function sendEmailOnSubmit(e) {
    try {
        // Extract values as per the columns in the sheet
        //var timestamp = e.values[0];
        var email = e.values[1];
        //var name = e.values[2];
        //var phone = e.values[3];


        if (!isValidEmail(email)) {
            throw new Error("Invalid email format.");
        }

        var subject = SUBJECT;

        // Pass all the required the values to the template function
        var body = getEmailTemplate(email);

        GmailApp.sendEmail(email, subject, "", {htmlBody: body});

    } catch (error) {
        logError(error, new Date(), e.values);
    }
}

function setupTrigger() {
    // Check if a trigger already exists
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === "sendEmailOnSubmit") {
            console.log("Trigger already exists.");
            return;
        }
    }

    // Create a trigger if it doesn't exist
    ScriptApp.newTrigger("sendEmailOnSubmit")
        .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
        .onFormSubmit()
        .create();
}
