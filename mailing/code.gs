// Sheets names
const SHEET_MAILS = "Mails";

// Function to convert number to char
function getChar(number) { return String.fromCharCode(65 + number); }
// Mails' sheet constants
const MAILS_MAIL_C = 1;
const MAILS_MAIL_COL = getChar(MAILS_MAIL_C);

// Spreadsheets
const SS_MAILS = SpreadsheetApp.getActive().getSheetByName(SHEET_MAILS);
// User Interface
const UI = SpreadsheetApp.getUi();
// Sheet
const SHEET = SpreadsheetApp.getActiveSheet();

// main function
function segmentate() {
    // Gets data
    var data = SS_MAILS.getDataRange().getValues();
    var lastHiddenRow = getLastHiddenRow(SS_MAILS, MAILS_MAIL_COL);
    var lastRow = getLastRow(SS_MAILS.getRange(getWholeColumnNotation(MAILS_MAIL_COL)))
    
    /*
      Reads mails
    */
    // A 'lastHiddenRow+1' is not needed because the DataRange numeration starts at 0,0
    for (var i = lastHiddenRow; i < lastRow; i++) {
      // gets whole mail
      var mail = data[i][MAILS_MAIL_C];
      atIndex = mail.indexOf('@'); // index of at symbol @

      // gets the number of boleta, which will help us segmantate
      var num = mail.substring(atIndex - 4, atIndex - 2);

      Logger.log(num);
    }
}

function segmentate(num, name, mail) {
  if (num < 10) {
    // Segmantate to GF
  }
  else if (10 <= num < 13) {
    // Segmantate to oGTa
  }
  else {
    // Segmantate to oGV
  }
}

// this function only works with contigous hidden rows
function getLastHiddenRow(sheet, columnL) {
  var lastHiddenRow = 1; // The row 0 does not count because it's the header
  // Last row, in certain column, with data in it will be the limit
  var lastRow = getLastRow(sheet.getRange(getWholeColumnNotation(columnL)));

  if (!sheet.isRowHiddenByUser(2))
    return lastHiddenRow;

  for (var i = 1; i < lastRow; i++) {
    if (sheet.isRowHiddenByUser(i + 1))
      lastHiddenRow++;
  }

  return lastHiddenRow;
}

/**
 * Gets the last row number within a given range
 * @param {Range} A range can be a single cell in a sheet or a group of adjacent cells in a sheet
 * @returns {number} The index of the last row with data in it
 */
function getLastRow(range) {
  return range.getValues().filter(String).length;
}

// 
function getWholeColumnNotation(columnL) {
  return columnL + ":" + columnL;
}