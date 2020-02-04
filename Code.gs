var SCRIPT_NAME = 'Email_Generator'
var SCRIPT_VERSION = 'v1.9.dev_ajr'

function onOpen() {
	var ui = SpreadsheetApp.getUi();

	ui.createMenu('CloudFire')
		.addItem('Generate HTML Email', 'generateHtmlEmail')
        .addSeparator()
        .addItem('Add New Fields for Input', 'addNewFieldsForInput')
        .addSeparator()
        .addItem('Hide Empty Rows', 'hideEmptyRows')
        .addItem('Show All Rows', 'showAllRows')
        .addSeparator()
        .addItem('Reformat Spreadsheet', 'reformatSpreadsheet')
        .addSeparator()
        .addItem('Hide Old Columns', 'hideOldColumns')
        .addItem('Show All Columns', 'showAllColumns')
        .addSeparator()
        .addItem('Delete Empty Columns', 'removeEmptyColumns')
		.addToUi();
}

function generateHtmlEmail()    {EmailGenerator.generateHtmlEmail()}
function addNewFieldsForInput() {EmailGenerator.addNewFieldsForInput()}
function hideEmptyRows()        {EmailGenerator.hideEmptyRows()}
function showAllRows()          {EmailGenerator.showAllRows()}
function reformatSpreadsheet()  {EmailGenerator.reformatSpreadsheet()}
function hideOldColumns()       {EmailGenerator.hideOldColumns()}
function showAllColumns()       {EmailGenerator.showAllColumns()}
function removeEmptyColumns()   {EmailGenerator.removeEmptyColumns()}
