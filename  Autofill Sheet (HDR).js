function fillSpecificColumnsHDR() {
    var sheetId = "1trq37tFuu6a0tCgrM9ZfcHsQvG4VqEmfwUYkjKmaZVQ"; // Replace with your actual sheet ID
    var sheetName = "Half Day Request"; // Replace with your actual sheet name
    var sheet = SpreadsheetApp.openById(sheetId).getSheetByName(sheetName);
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  // Define the header row index and the columns we are interested in
  var requestNoIndexCol = 0; // Column A (index 0)
  var totalPaidLeavesCol = 4; // Column E (index 4)
  var totalLeavesTakenCol = 5; // Column F (index 5)
  var totalRequestRaised = 6; // Column G (index 6)
  var statusCol = 8; // Column I (index 8)
  var totalLeavesRequested = 14; // Column O (index 14)

  // Loop through each row starting from the second row
  for (var row = 1; row < values.length; row++) {
    rangecounter=row+1
    values[row][requestNoIndexCol] = row ;
    values[row][totalPaidLeavesCol] = "= FILTER( 'Employee Data'!I3:I1000 ,'Employee Data'!B3:B1000 = C"+rangecounter+")" ;
    values[row][totalLeavesTakenCol] = "= FILTER( 'Employee Data'!J3:J1000 ,'Employee Data'!B3:B1000 = C"+rangecounter+")" ;
    values[row][totalRequestRaised] = "=COUNTIF(C2:C"+rangecounter+",C"+rangecounter+")";
    
    values[row][totalLeavesRequested] = "1" ;


  }

  // Set the updated values back to the range
  range.setValues(values);
}


// Completed and Working 
// Automation by 1 Minute