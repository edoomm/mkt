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

/**
 * Main function that is associated with the button of the SpreadSheet
 */
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

/**
 * Auxiliar function that will segmentate mails depending on the year the student came in
 * @param num The two digit number needed to know in which year a student entered IPN
 * @param name The name of the student
 * @param mail The institutional mail
 */
function segmentate(num, name, mail) {
  if (num < 10) {
    // TODO: Segmantate to GF
  }
  else if (10 <= num < 13) {
    // TODO: Segmantate to oGTa
  }
  else {
    // TODO: Segmantate to oGV
  }
}

/**
 * Gets the last hidden row by user in a certain column.
 * This function only works with contigous hidden rows, i.e.: Rows should go from 2-15, you cannot use it with different ranges like 2-4, 7-10, etc
 * @param   {SpreadsheetApp.Sheet}  sheet   The sheet where the row index will retrieved
 * @param   {string}                columnL The letter of the column to analyze
 * @returns {number}                        The index of the last hidden row by user
 */
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
 * @param   {SpreadsheetApp.Range}   range A range can be a single cell in a sheet or a group of adjacent cells in a sheet
 * @returns {number}                       The index of the last row with data in it
 */
function getLastRow(range) {
  return range.getValues().filter(String).length;
}

/**
 * Gets the formatted string to have a whole column
 * @param   {string}  columnL A column letter
 * @returns {string}          A formatted string based on the column letter, i.e.: "A:A"
 */
function getWholeColumnNotation(columnL) {
  return columnL + ":" + columnL;
}