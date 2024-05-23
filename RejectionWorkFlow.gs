function sendrejectLeaveStatusEmailById() {
  var sheetId = '1538283278';  // Replace with your sheet ID
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

  // Iterate through the rows
  for (var i = 2; i < data.length; i++) {
    var row = data[i];
    var status = row[statusIndex];
    var intimationSent = row[intimationSentIndex];
   

    // Check if status is present and email has not been sent
    if (status === "Rejected" && intimationSent !== true) {
      var requestNo = row[requestNoIndex];
      var employeeName = row[employeeNameIndex];
      var employeeEmail = row[employeeEmailIndex];
      
      var fromDate = row[fromDateIndex];
      var tillDate = row[tillDateIndex];
      var reasonForLeave = row[reasonForLeaveIndex];
      var personAtAuthority = row[personAtAuthorityIndex];

      var subject = 'Leave Request Status Update';
      var message = 'Dear ' + employeeName + ',\n\n' +
                    'Your leave request (' + requestNo + ') status has been updated.\n\n' +
                    'Leave Dates: From ' + fromDate + ' To ' + tillDate + '\n' +
                    'Reason: ' + reasonForLeave + '\n' +
                    'Status: ' + status + '\n' +
                    'Reviewed by: ' + personAtAuthority + '\n\n' +
                    'Best regards,\nYour Company';

      MailApp.sendEmail(employeeEmail, subject, message);

      // Mark as email sent by setting "TRUE" in the "Intimation Sent" column
      sheet.getRange(i + 1, intimationSentIndex + 1).setValue('TRUE');
    }
  }
}

function createTimeDrivenTrigger() {
  ScriptApp.newTrigger('sendrejectLeaveStatusEmailById')
      .timeBased()
      .everyMinutes(5) // Adjust the interval as needed
      .create();
}

// Schedule the function to run periodically
function createTimeDrivenTrigger() {
  ScriptApp.newTrigger('sendrejectLeaveStatusEmailById')
      .timeBased()
      .everyMinutes(1) // Adjust the interval as needed
      .create();
}
