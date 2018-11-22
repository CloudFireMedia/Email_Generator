// redevelopment note: only hide the column if the date in the first Row has lapsed

// move column D to the end of the sheet and hide onOpen
function moveColumn(iniCol, finCol) {  

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  var iniCol = 3;
  var finCol = sh.getMaxColumns();
  var lRow = sh.getMaxRows();

    sh.insertColumnAfter(finCol);
    var iniRange = sh.getRange(1, iniCol + 1, lRow);
    var finRange = sh.getRange(1, finCol + 1, lRow);
    iniRange.copyTo(finRange, {contentsOnly:true});
    sh.deleteColumn(iniCol + 1);    
    sh.hideColumns(finCol);
    sh.insertColumns(4, 1);
}