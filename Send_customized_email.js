// This constant is written in column ? for rows for which an email
// has been sent successfully.
var EMAIL_SENT = 'EMAIL_SENT';

function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 41; // First row of data to process
  var numRows = 42; // Number of rows to process //last_row - first_row + 1
  var startCol = 2; //First Column of data to process
  var numCols = 18; // Number of columns to process
  // Fetch the range of cells B2:C311
  var dataRange = sheet.getRange(startRow, startCol, numRows, numCols); //getRange(starting-row, starting-column, numRows, numCols) indexing starts with 1.
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  var TemplateText = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet2Template").getRange(1,1).getValue();
  
  for (var i = 0; i < data.length; ++i) {
    var col = data[i];
    var message = TemplateText.replace("{SPL}", col[0]).replace("{Full Name}", col[1]).replace("{Selected Council}", col[2]).replace("{Email Address}", col[3]).replace("{Current Institution}", col[5]).replace("{Phone Number}", col[6]).replace("{Your Delegation}", col[8]).replace("{Registered time slot}", col[16]);
    var emailAddress = col[3];
    var Status = col[17]; // row[numcols - 1] //The dec index column of status emailsent in created table range //data[i][4]
    var subject = '[HMUN’20] INTERVIEW SLOT REGISTRATION CONFIRMATION';
 
   
    if (Status !== EMAIL_SENT) { // Prevents sending duplicates
      GmailApp.sendEmail(emailAddress, subject, message, {
        name: "Hanoi Model UN"});
      sheet.getRange(startRow + i, 19).setValue(EMAIL_SENT); //number 3 is the dec value of status column starting with A
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush(); 
    } 
  }
}