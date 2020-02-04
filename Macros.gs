function addNewFieldsForInput() {
  var spreadsheet = SpreadsheetApp.getActive();
  var numColumns = spreadsheet.getLastColumn()-5;
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Input'), true);
  spreadsheet.getActiveSheet().insertColumnsBefore(6, 1);
  spreadsheet.getRange('Format!F:F').copyTo(spreadsheet.getRange('Input!F:F'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  hideOldColumns();
  spreadsheet.getRange('F2').activate();
};


function hideEmptyRows() {
  var s = SpreadsheetApp.getActive().getSheetByName('Input');
  s.showRows(1, s.getMaxRows());
  s.getRange('F:F')
  .getValues()
  .forEach( function (r, i) {
    if (r[0] == '') 
      s.hideRows(i + 1);
  });
}

function showAllRows() {
  var s = SpreadsheetApp.getActive().getSheetByName('Input');
  var rRows = s.getRange("A:A");
  s.unhideRow(rRows);
}


function reformatSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var numColumns = sheet.getMaxColumns();  
  var numRows = sheet.getMaxRows();  
  sheet.getRange(1,1,numRows,numColumns).clearFormat();
  ss.getRange('Format!A:F').copyTo(sheet.getRange(1,1,1,6), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  ss.getRange('Format!G:G').copyTo(sheet.getRange(1,7,numRows,numColumns-6), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
}



function hideOldColumns() {
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

function showAllColumns() {
  var spreadsheet = SpreadsheetApp.getActive();
  var numColumns = spreadsheet.getLastColumn();
  //  spreadsheet.getRange('F:F').activate();
  spreadsheet.getActiveSheet().showColumns(1, numColumns);
  spreadsheet.getRange('F2').activate();
};


function removeEmptyColumns() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('Input');
  var numColumns = spreadsheet.getLastColumn();  
  for (var i = numColumns - 1; i>=6; i--) {
    if (sheet.getRange(1,i+1,sheet.getMaxRows(),1).isBlank()) {
      sheet.deleteColumn(i+1); 
      Logger.log(i+1)
    }  
  }
}