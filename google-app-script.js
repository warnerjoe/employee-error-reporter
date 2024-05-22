// Creates custom menu to execute scripts from
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Scripts')
    .addItem('Copy Rows', 'copyRowsToEmployeeTabs')
    .addItem('Clear Entry Sheet', 'clearEntrySheet')
    .addSeparator()
    .addItem('Unhide All Sheets', 'unhideAllSheets')
    .addItem('Hide All Sheets', 'hideAllSheetsExceptFirst')
    .addSeparator()
    .addItem('Sort Employees', 'sortEmployeeColumn')
    .addToUi();
}

// Takes all new entries and copies their contents to the proper employee sheet.
function copyRowsToEmployeeTabs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var entrySheet = getSheetByName('Entry');
  if (entrySheet) {
    var data = entrySheet.getRange(2,1, entrySheet.getLastRow() - 1, entrySheet.getLastColumn()).getValues();
    data.forEach(function(row) {
      var employee = row[0];
      if (employee) {
        var employeeSheet = getOrCreateSheet(ss, employee);
        if (!rowExists(employeeSheet, row)) {
          employeeSheet.appendRow(row);
        }
      }
    });
  }
}

// Checks to see if the row already exists in the target sheet, to prevent duplicate entries
function rowExists(sheet, row) {
  var data = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  return data.some(function(existingRow) {
    return existingRow.join() === row.join();
  });
}

// Locates the proper sheet using the name from Column 1 of the entry form
function getSheetByName(sheetName) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
}

// If the sheet does not already exist, it creates one.  Still needs work.
function getOrCreateSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

// Sorts the employee column of the Entry sheet.
function sortEmployeeColumn() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entry');
  var range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  range.sort({column: 1, ascending: true});
}

// Clears the entry sheet after the contents have been copied.
function clearEntrySheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Entry');

  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
}

// Unhides all of the employee sheets.
function unhideAllSheets() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();

  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    if (sheet.isSheetHidden()) {
      sheet.showSheet();
    }
  }
}

// Hides all the employee sheets, leaving only the Entry sheet.
function hideAllSheetsExceptFirst() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();

  if (sheets.length > 0) {
    for (var i = 1; i < sheets.length; i++) {
      var sheet = sheets[i];
      sheet.hideSheet();
    }
  }
}