function autorejectHDR() {
    var sheetId = "1trq37tFuu6a0tCgrM9ZfcHsQvG4VqEmfwUYkjKmaZVQ"; // Replace with your actual sheet ID
    var sheetName = "Half Day Request"; // Replace with your actual sheet name
    var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
    var range = sheet.getDataRange();
    var values = range.getValues();
  
    var timestampColumn = 7; // Column H, converted to 0-indexed
    var startdateColumn = 9; // Column J, converted to 0-indexed
    var statusColumn = 8; // Column I, converted to 0-indexed
  
    // Loop through each row starting from the second row (index 1)
    for (var row = 1; row < values.length; row++) {
      var requestTimestamp = new Date(values[row][timestampColumn]);
      var leaveStartDate = new Date(values[row][startdateColumn]);
      // Logger.log(isSameDate(requestTimestamp, leaveStartDate));
      // Logger.log(requestTimestamp.getHours() >= 9);
      // Compare request time and leave start date
      if (requestTimestamp.getHours() >= 9 && isSameDate(requestTimestamp, leaveStartDate)) {
  
        // Set status to "Rejected"
        values[row][statusColumn] = "Rejected";
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
  