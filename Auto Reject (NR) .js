function autorejectNR() {
    var sheetId = "1trq37tFuu6a0tCgrM9ZfcHsQvG4VqEmfwUYkjKmaZVQ"; // Replace with your actual sheet ID
    var sheetName = "New Request"; // Replace with your actual sheet name
    var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
    var range = sheet.getDataRange();
    var values = range.getValues();
  
  
    var applicantColumn= 1 ;
    var reasonColumn = 11 ;
    var timestampColumn = 7; // Column H, converted to 0-indexed
    var startdateColumn = 9; // Column J, converted to 0-indexed
    var statusColumn = 8; // Column I, converted to 0-indexed
    var personAtAuthority= 12 ; // Column M , convereted to 0-indexed
    var remarks = 15 ; //Column Q , converted to 0-indexed
    Email = "leavesystem.fls@gmail.com"
  
    // Loop through each row starting from the second row (index 1)
    for (var row = 1; row < values.length; row++) {
      var requestTimestamp = new Date(values[row][timestampColumn]);
      var leaveStartDate = new Date(values[row][startdateColumn]);
      var resend = values[row][personAtAuthority];
      // Logger.log(isSameDate(requestTimestamp, leaveStartDate));
      // Logger.log(requestTimestamp.getHours() >= 9);
      // Compare request time and leave start date
      Logger.log(resend);
  
      if (requestTimestamp.getHours() >= 9 && isSameDate(requestTimestamp, leaveStartDate) && resend !== "System Rejected" ) {
  
        // Set status to "Rejected"
        values[row][statusColumn] = "Rejected";
        values[row][personAtAuthority] = "System Rejected";
        values[row][remarks]= "Leave Applied After 9:00 AM. Please contact mentor if urgent."
        reasonForLeave = values[row][reasonColumn];
        applicant = values[row][applicantColumn];
        
  
        var subject = 'Leave Request Auto Rejected';
        var message = 'Dear Mentors,\n ' +
                      'A leave request has been Auto Rejected.\n\n' +
                      'Leave Applicant:  ' +  applicant +'\n'+ 
                      
                      'Reason: ' + reasonForLeave + '\n' +
                    
                      'Best regards,\nLeave Management System';
  
        MailApp.sendEmail(Email, subject, message);
  
      }
    }
  
    // Set the updated values back to the range
    range.setValues(values);
  }
  
  // Helper function to check if two dates are on the same day
  function isSameDate(date1, date2) {
    return date1.getFullYear() === date2.getFullYear() &&
           date1.getMonth() === date2.getMonth() &&
           date1.getDate() === date2.getDate();
  }
  