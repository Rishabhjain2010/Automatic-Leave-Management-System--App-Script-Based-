function sendrejectLeaveStatusEmailByIdCR() {
    var sheetId = '1369831212';  // Replace with your sheet ID
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
    var employeeNameIndex = headers.indexOf("Employee Name");
    var employeeEmailIndex = headers.indexOf("Employee Email");
    var statusIndex = 8;
    // Logger.log(statusIndex);
    var fromDateIndex = 9;
    var tillDateIndex = 10;
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
  
        var subject = 'Leave Request Cancellation Status Update';
        var message = 'Dear ' + employeeName + ',\n\n' +
                      'Your leave request (' + requestNo + ') status has been updated.\n\n' +
                      'Leave Dates: From ' + fromDate + ' To ' + tillDate + '\n' +
                      'Reason: ' + reasonForLeave + '\n' +
                      'Status: ' + status + '\n' +
                    
                      'Best regards,\nFostering Linux Services Pvt. Ltd.';
  
        MailApp.sendEmail(employeeEmail, subject, message);
  
        // Mark as email sent by setting "TRUE" in the "Intimation Sent" column
        sheet.getRange(i + 1, intimationSentIndex + 1).setValue('TRUE');
      }
    }
  }
  