function sendLeaveStatusEmailById() {
  var sheetId = '505062935';  // Replace with your sheet ID
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheets().filter(function(sh) {
    return sh.getSheetId() == sheetId;
  })[0];

  if (!sheet) {
    Logger.log('Sheet not found');
    return;
  }
  
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();

  // Assuming the first row is the header
  var headers = data[0];
  
  // Get column indices
  var requestNoIndex = headers.indexOf("Request No.");
  var employeeNameIndex = headers.indexOf("Employee Name");
  var employeeEmailIndex = headers.indexOf("Employee EMail");
  var statusIndex = headers.indexOf("Status");
  var fromDateIndex = headers.indexOf("Leave Dates From");
  var tillDateIndex = headers.indexOf("Leave Dates Till");
  var reasonForLeaveIndex = headers.indexOf("Reason for Leave");
  var personAtAuthorityIndex = headers.indexOf("Person at Authority");
  var intimationSentIndex = headers.indexOf("Intimation Sent");

  // Get your calendar by its ID
  var myCalendarId = 'leavesystem.fls@gmail.com'; // Replace with your own email ID or calendar ID
  var calendar = CalendarApp.getCalendarById(myCalendarId);
  
  if (!calendar) {
    Logger.log('Calendar not found');
    return;
  }

  // Iterate through the rows
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var status = row[statusIndex];
    var intimationSent = row[intimationSentIndex];


    // Check if status is "Approved" and email has not been sent
    if (status === "Approved" && intimationSent !== true) {
      var requestNo = row[requestNoIndex];
      var employeeName = row[employeeNameIndex];
      var employeeEmail = row[employeeEmailIndex];
      var fromDate = new Date(row[fromDateIndex]);
      var tillDate = new Date(row[tillDateIndex]);
      var reasonForLeave = row[reasonForLeaveIndex];
      var personAtAuthority = row[personAtAuthorityIndex];

      Logger.log('Processing request no: ' + requestNo);
      Logger.log('From Date: ' + fromDate + ', Till Date: ' + tillDate);

      // Check date validity
      if (isNaN(fromDate.getTime()) || isNaN(tillDate.getTime())) {
        Logger.log('Invalid dates for request no: ' + requestNo);
        continue;
      }

      var subject = 'Leave Request Approval Notification';
      var message = 'Dear ' + employeeName + ',\n\n' +
                    'Your leave request (' + requestNo + ') has been approved.\n\n' +
                    'Leave Dates: From ' + fromDate.toDateString() + ' To ' + tillDate.toDateString() + '\n' +
                    'Reason: ' + reasonForLeave + '\n' +
                    'Reviewed by: ' + personAtAuthority + '\n\n' +
                    'Best regards,\nYour Company';

      try {
        MailApp.sendEmail(employeeEmail, subject, message);

        // Add event to your Google Calendar
        calendar.createEvent(employeeName + ' - Leave (' + reasonForLeave + ')',
                            fromDate,
                            tillDate,
                            {description: 'Approved leave request for ' + reasonForLeave});
        Logger.log('Event created successfully for request no: ' + requestNo);

        // Mark as email sent by setting "TRUE" in the "Intimation Sent" column
        sheet.getRange(i + 1, intimationSentIndex + 1).setValue('TRUE');
        
        // Pause to avoid hitting email rate limits
        Utilities.sleep(1000); // Sleep for 1 second
      } catch (e) {
        Logger.log('Error for request no: ' + requestNo + ': ' + e.message);

        // Break the loop if too many emails sent in a day to avoid hitting the limit
        if (e.message.includes('Service invoked too many times for one day')) {
          break;
        }
      }
    }
  }
}

// Schedule the function to run periodically
function createTimeDrivenTrigger() {
  ScriptApp.newTrigger('sendLeaveStatusEmailById')
      .timeBased()
      .everyMinutes(1) // Adjust the interval as needed
      .create();
}

