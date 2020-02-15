function addNewFieldsForInput_() {
  var spreadsheet = SpreadsheetApp.getActive();
  var numColumns = spreadsheet.getLastColumn() - 5;
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Input'), true);
  spreadsheet.getActiveSheet().insertColumnsBefore(6, 1);
  spreadsheet.getRange('Format!F:F').copyTo(spreadsheet.getRange('Input!F:F'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  hideOldColumns_();
  spreadsheet.getRange('F2').activate();
  Log_.info('New fields added.')
};

function hideEmptyRows_() {
  var s = SpreadsheetApp.getActive().getSheetByName('Input');
  s.showRows(1, s.getMaxRows());
  s.getRange('F:F').getValues().forEach(function (r, i) {
    if (r[0] === '') s.hideRows(i + 1);
  });
}

function showAllRows_() {
  var s = SpreadsheetApp.getActive().getSheetByName('Input');
  var rRows = s.getRange("A:A");
  s.unhideRow(rRows);
}

function reformatSpreadsheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0]; // "Input"
  var numColumns = sheet.getMaxColumns();  
  var numRows = sheet.getMaxRows();  
  sheet.getRange(1,1,numRows,numColumns).clearFormat();
  ss.getRange('Format!A:F').copyTo(sheet.getRange(1,1,1,6), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  ss.getRange('Format!G:G').copyTo(sheet.getRange(1,7,numRows,numColumns - 6), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  Log_.info('Input tab reformatted');
}

function hideOldColumns_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var numColumns = sheet.getLastColumn();
  var v = sheet.getRange(1,1,1,numColumns).getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  for (var i = sheet.getLastColumn(); i > 6; i--) {
    var t = v[0][i - 1];
    var u = new Date(t);
    if ((u < today) || (typeof t === 'string' || t instanceof String)) { 
      sheet.hideColumns(i);
    }
  }
}

function showAllColumns_() {
  var spreadsheet = SpreadsheetApp.getActive();
  var numColumns = spreadsheet.getLastColumn();
  spreadsheet.getActiveSheet().showColumns(1, numColumns);
  spreadsheet.getRange('F2').activate();
}

function removeEmptyColumns_() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('Input');
  var numColumns = spreadsheet.getLastColumn();  
  for (var i = numColumns - 1; i>=6; i--) {
    if (sheet.getRange(1,i+1,sheet.getMaxRows(),1).isBlank()) {
      sheet.deleteColumn(i+1); 
      Log_.info('Deleted column: ' + i + 1)
    }  
  }
}

function archiveCurrentColumn_() {
  var spreadsheet = SpreadsheetApp.getActive();
  var activeSheet = spreadsheet.getActiveSheet();
  var activeRange = spreadsheet.getActiveRange();
  var activeColumnNumber = activeRange.getColumn();
  var numberOfRows = activeSheet.getLastRow();
  var activeColumnRange = activeSheet.getRange(1, activeColumnNumber, numberOfRows, 1);
  var activeColumnValues = activeColumnRange.getValues()  
  var numberOfColumns = activeSheet.getLastColumn();
  activeSheet.insertColumnAfter(numberOfColumns);
  activeSheet.getRange(1, numberOfColumns + 1, numberOfRows, 1).setValues(activeColumnValues);
  activeSheet.hideColumns(numberOfColumns + 1);
  Log_.info('Added the values in the active column to the end of the columns and hid it')

  activeColumnRange.clear();
  spreadsheet.getRange('Format!F:F').copyTo(activeColumnRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  activeSheet.getRange(1, activeColumnNumber).setValue('YYYY.MM.DD');
  activeSheet.getRange(3, activeColumnNumber).setValue('Email Subject');   
  Log_.info('Cleared the contents of the active column and pasted in the default format')
}