function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Gmail Tools')
      .addItem('Get Emails', 'getAllSenderNames')
    // .addItem('Clear Data', 'clearSheet')
    .addItem('Delete Emails', 'deleteEmails')
    .addToUi();
}

function sortMessages() {
  // Implement sorting logic here
}
function deleteEmails() {
  // Get the active spreadsheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the data range
  var range = sheet.getDataRange();
  
  // Get all values in the range
  var values = range.getValues();
  
  // Loop through each row
  for (var i = 0; i < values.length; i++) {
    // If the checkbox is checked in the current row
    if (values[i][1] === true) { // Assuming checkboxes are in column B (index 1)
      // Get the sender name from column A
      var senderName = values[i][0]; // Assuming sender names are in column A (index 0)
      
      // Get all threads in the inbox
      var threads = GmailApp.search('from:' + senderName);
      
      // Delete all emails in the threads
      for (var j = 0; j < threads.length; j++) {
        threads[j].moveToTrash();
      }
       sheet.getRange(i + 1, 2).setValue(false);
    }
  }
  getAllSenderNames();
}
function getAllSenderNames() {
  // Get the active spreadsheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

var lastRow = sheet.getLastRow();
 

  if (lastRow == 0) {
    // Run the code to create data
   
   createData();
    deleteEmptyRows();
  } else {
    // Run the clear function
    clearSheet();
   
  }


  // Get all threads in the inbox
 
}

function createData(){

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get the data range
  var range = sheet.getDataRange();
   var threads = GmailApp.getInboxThreads();

  // Create an object to store sender names
  var senderNames = {};

  // Loop through each thread
  for (var i = 0; i < threads.length; i++) {
    // Get the messages in the current thread
    var messages = threads[i].getMessages();

    // Loop through each message in the current thread
    for (var j = 0; j < messages.length; j++) {
      // Get the sender of the current message
      var sender = messages[j].getFrom();

      // Extract the sender name from the sender email address
      var senderName = sender.substring(0, sender.indexOf('<')).trim();

      // Add the sender name to the senderNames object
      if (!senderNames[senderName]) {
        senderNames[senderName] = true;
      }
    }
  }

  // Convert the senderNames object keys to an array
  var uniqueSenderNames = Object.keys(senderNames);

  // Write sender names to the spreadsheet
  for (var k = 0; k < uniqueSenderNames.length; k++) {
    // Skip empty sender names
    if (uniqueSenderNames[k] !== "") {
      sheet.getRange('A' + (k + 1)).setValue(uniqueSenderNames[k]);
    }
  }

  // Insert checkboxes in column B for each row
  if (uniqueSenderNames.length > 0) {
    sheet.getRange('B1:B' + uniqueSenderNames.length).insertCheckboxes();
  }
}
function clearSheet() {
   var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Get the last row with data in column A (assuming sender names are in column A)
  var lastRow = sheet.getLastRow();

  // Get the range for columns A and B
  var rangeA = sheet.getRange('A1:A' + lastRow);
  var rangeB = sheet.getRange('B1:B' + lastRow);

  // Clear content and formatting for both ranges
  rangeA.clear({contentsOnly: true, skipFilteredRows: true});
  rangeB.clear({contentsOnly: true, skipFilteredRows: true});
}

function deleteEmptyRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();

  // Iterate over the rows from the last row to the first row
  for (var i = lastRow; i > 0; i--) {
    // Check if the value in column A of the current row is empty
    if (sheet.getRange('A' + i).isBlank()) {
      // If the cell is empty, delete the entire row
      sheet.deleteRow(i);
    }
  }
}

