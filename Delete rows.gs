function deleteRows() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Responses");

var start, end;

start = 2;
end = sheet.getLastRow() - 1;

sheet.deleteRows(start, end);
}