//Main function to be triggered based on time
//This function only runs on Google's App Script Platform
//

function readEmails() {
    var threads = GmailApp.search('in:inbox is:unread subject:"Leave Request"');
    var sheetId = "1trq37tFuu6a0tCgrM9ZfcHsQvG4VqEmfwUYkjKmaZVQ"; // Replace with your actual sheet ID
    var sheetName = "New Requests"; // Replace with your actual sheet name
    var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
    var lastRow = sheet.getLastRow();  // Get the last row with data

    // Check if the sheet is empty
    if (lastRow === 0) {
        // Append header row if the sheet is empty
        sheet.appendRow(["Request No", "Employee Name", "Employee ID", "Email Address", "Total Paid Leaves", "Total Leaves Taken", "Remaining Leaves", "Request Date", "Status", "Start Date", "End Date", "Reason", "Person at Authority", "Attachment", "Attachment Link", "Total Leaves Requested"]);
    } else if (lastRow > 0) {
        lastRow++; // Increment lastRow to append data starting from the next row
    }

    var requestNo = lastRow-2 ;  // Increment the request number from the last row (-2 for error handling)

    threads.forEach(thread => {
        var messages = thread.getMessages();
        messages.forEach(message => {
            var emailBody = message.getPlainBody(); // Use getPlainBody() instead of getBody()
            var emailAddress = message.getFrom();
            var requestDate = message.getDate();
            var details = parseEmail(emailBody);

            // Initialize attachment link variables
            var attachmentLinks = [];

            // Check for attachments
            var attachments = message.getAttachments();
            for (var i = 0; i < attachments.length; i++) {
                // Get the message ID
                var messageId = message.getId();
                // Construct the link to the attachment using the index
                var attachmentLink = "https://mail.google.com/mail/u/0/?view=att&th=" + messageId + "&attid=" + i;
                attachmentLinks.push(attachmentLink); // Store the attachment link
            }

            // Join attachment links into a single string
            var attachmentLinkStr = attachmentLinks.join(", ");
            var filterCounter=requestNo+2;
            sheet.appendRow([
                requestNo,
                details.name,
                "=FILTER('Employee Data'!C:C, 'Employee Data'!B:B = B" + filterCounter + ")" ,
                emailAddress,
                "",  // Total Paid Leaves - to be filled manually or through another system
                "",  // Total Leaves Taken - to be filled manually or through another system
                "",  // Remaining Leaves - to be filled manually or through another system
                requestDate,
                "Pending",
                details.startDate,
                details.endDate,
                details.reason,
                "",  // Person at Authority - to be filled manually or through another system
                
                attachmentLinkStr,  // Attachment links
                details.totalLeavesRequested
            ]);

            message.markRead();
            requestNo++;
        });
    });
}





function parseEmail(emailBody) {
    var lines = emailBody.split('\n');
    var details = {};
    
    lines.forEach((line, index) => {
        Logger.log("Line " + index + ": " + line);
        if (line.includes('Employee Name')) {
            details.name = line.split(':')[1].trim();
        } else if (line.includes('Employee ID')) {
            details.id = line.split(':')[1].trim();
        } else if (line.includes('From')) {
            Logger.log("Start Date Line: " + line);
            details.startDate = line.split(':')[1].trim();
        } else if (line.includes('Till')) {
            Logger.log("End Date Line: " + line);
            details.endDate = line.split(':')[1].trim();
        } else if (line.includes('Reason')) {
            details.reason = line.split(':')[1].trim();
        }
    });

    // Calculate total leaves requested
    var startDate = new Date(details.startDate.split('-').reverse().join('-'));
    var endDate = new Date(details.endDate.split('-').reverse().join('-'));
    var totalLeavesRequested = Math.ceil((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
    details.totalLeavesRequested = totalLeavesRequested;

    return details;
}



function sendNotification(email, status) {
    var subject = "Your Leave Request Status Update";
    var body = "Your leave request has been " + status + ".";
    MailApp.sendEmail(email, subject, body);
}

function addEventToCalendar(details) {
    var calendar = CalendarApp.getCalendarById("your-calendar-id@group.calendar.google.com");
    var startDate = new Date(details.startDate.split('-').reverse().join('-'));
    var endDate = new Date(details.endDate.split('-').reverse().join('-'));
    endDate.setDate(endDate.getDate() + 1); // Include the end date as a full day
    var event = calendar.createEvent(details.reason, startDate, endDate);
    event.setDescription("Leave approved for " + details.name);
}

function updateLeaveStatus(row, status) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("New Request");
    var email = sheet.getRange(row, 4).getValue(); // Employee Email is in the 4th column
    sendNotification(email, status);
    sheet.getRange(row, 9).setValue(status); // Status is in the 9th column
    if (status === "Approved") {
        var details = {
            name: sheet.getRange(row, 2).getValue(),
            startDate: sheet.getRange(row, 10).getValue(), // Leave Dates From
            endDate: sheet.getRange(row, 11).getValue(),   // Leave Dates Till
            reason: sheet.getRange(row, 12).getValue()     // Reason for Leave
        };
        addEventToCalendar(details);
    }
}

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Leave Management')
      .addItem('Approve Leave', 'approveLeave')
      .addItem('Deny Leave', 'denyLeave')
      .addToUi();
}

function approveLeave() {
    var sheetId = "1trq37tFuu6a0tCgrM9ZfcHsQvG4VqEmfwUYkjKmaZVQ"; // Replace with your actual sheet ID
    var sheetName = "New Requests"; // Replace with your actual sheet name
    var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
    var range = sheet.getActiveRange();
    var row = range.getRow();
    updateLeaveStatus(row, "Approved");
}

function denyLeave() {
    var sheetId = "1trq37tFuu6a0tCgrM9ZfcHsQvG4VqEmfwUYkjKmaZVQ"; // Replace with your actual sheet ID
    var sheetName = "New Requests"; // Replace with your actual sheet name
    var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
    var range = sheet.getActiveRange();
    var row = range.getRow();
    updateLeaveStatus(row, "Denied");
}


function readEmail() {
    var threads = GmailApp.search('in:inbox is:unread subject:"Leave Request"');
    threads.forEach(thread => {
        var messages = thread.getMessages();
        messages.forEach(message => {
            var attachments = message.getAttachments();
            Logger.log("Attachments for message with subject: " + message.getSubject());
            attachments.forEach(attachment => {
                Logger.log("Attachment Name: " + attachment.getName());
                Logger.log("Attachment Content Type: " + attachment.getContentType());
                Logger.log("Attachment Size: " + attachment.getSize());
                Logger.log("Attachment Keys: " + Object.keys(attachment));
            });
        });
    });
}

