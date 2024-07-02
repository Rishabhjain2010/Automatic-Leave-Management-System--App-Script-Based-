function sendRequestUpdateCR() {
    var sheetId = '682677884';  // Replace with your sheet ID
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
    var requestNoIndex = 0;
    var employeeNameIndex = headers.indexOf("Intern Name");
    var employeeEmailIndex = headers.indexOf("Intern Email");
    var statusIndex = headers.indexOf("Status");
    var fromDateIndex = headers.indexOf("Leave Dates From");
    var tillDateIndex = headers.indexOf("Leave Dates Till");
    var reasonForLeaveIndex = headers.indexOf("Reason for Leave");
    var personAtAuthorityIndex = headers.indexOf("Person at Authority");
    var intimationSentIndex = headers.indexOf("Intimation Sent");
    
    // Iterate through the rows
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var status = row[statusIndex];
      var intimationSent = row[intimationSentIndex];
      
  
      // Check if status is present and email has not been sent
      if (intimationSent !== true ) {
        var requestno = row[requestNoIndex]
        var employeeName = row[employeeNameIndex];
        var employeeEmail = row[employeeEmailIndex];
        var fromDate = row[fromDateIndex];
        var tillDate = row[tillDateIndex];
        var reasonForLeave = row[reasonForLeaveIndex];
        var subject = 'Cancellation Leave Request Registered';
        var message = 'Dear ' + employeeName + ',\n\n' +
                      'Your leave cancellation request has been received and Registered.\n\n' +
                      'Your Request No. is' + requestno +'\n'+ 
                      'Leave Dates: From ' + fromDate + 'Till:' + tillDate + ' \n' +
                      'Reason: ' + reasonForLeave + '\n' +
                    
                      'Best regards,\nFOSTERing Linux Services Pvt. Ltd.';
  
        MailApp.sendEmail(employeeEmail, subject, message);
  
        // Mark as email sent by setting "TRUE" in the "Intimation Sent" column
        sheet.getRange(i + 1, intimationSentIndex + 1).setValue('TRUE');
      }
    }
  }
  