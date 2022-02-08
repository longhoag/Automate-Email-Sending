// This constant is written in column ? for rows for which an email
// has been sent successfully.
var EMAIL_SENT = 'EMAIL_SENT';

function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 305; // First row of data to process
  var numRows = 7; // Number of rows to process //last_row - first_row + 1
  var startCol = 2; //First Column of data to process
  var numCols = 2; // Number of columns to process
  // Fetch the range of cells B2:C311
  var dataRange = sheet.getRange(startRow, startCol, numRows, numCols); //getRange(starting-row, starting-column, numRows, numCols) indexing starts with 1.
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  
  
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[0]; // First column in created table range  //data[i][0]
    var message = HtmlService.createTemplateFromFile('Letter').evaluate().getContent();
    var Status = row[1]; // row[numcols - 1] //The dec index column of status emailsent in created table range //data[i][4]
    var subject = '[HMUN 2020] EARLY DECISION REGISTRATION FORM IS OPENED!';
    
    if (Status !== EMAIL_SENT) { // Prevents sending duplicates
      GmailApp.sendEmail(emailAddress, subject, message, {
        name: "Hanoi Model UN",
        htmlBody: message});
      sheet.getRange(startRow + i, 3).setValue(EMAIL_SENT); //number 3 is the dec value of status column starting with A
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush(); 
    } 
  }
}