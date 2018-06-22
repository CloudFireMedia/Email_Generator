function deleteRows() {
	var ss = SpreadsheetApp.getActive(),
		sheet = ss.getSheetByName('Responses');

	sheet.deleteRows(2, (sheet.getLastRow() - 1));
}