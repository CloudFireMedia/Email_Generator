function onOpen() {
	var ui = SpreadsheetApp.getUi();

	ui.createMenu('[ Custom Menu ]')
		.addItem('Create HTML', 'createHtmlFromSelectedRow')
		.addItem('Set Defaults', 'setDefaults')
		.addSeparator()
		.addItem('Delete All Rows', 'deleteRows')
		.addToUi();
}

function createHtmlFromSelectedRow() {
	var ss = SpreadsheetApp.getActiveSpreadsheet(),
		sheet = ss.getSheetByName('Email Content'),
		data = sheet.getRange('D1:D18').getValues();

	onFormSubmit({
		'values': [
			'',
			data[0][0],
			data[1][0],
			data[2][0],
			data[3][0],
			data[4][0],
			data[5][0],
			data[6][0],
			data[7][0],
			data[8][0],
			data[9][0],
			data[10][0],
			data[11][0],
			data[12][0],
			data[13][0],
			data[14][0],
			data[15][0],
			data[16][0],
			data[17][0]
		]
	});

	ss.toast('Successfully processed!', 'Info', 2);
}

function setDefaults() {
	var ss = SpreadsheetApp.getActiveSpreadsheet(),
		sheet = ss.getSheetByName('Default'),
		data = sheet.getDataRange().getValues(),
		html = HtmlService.createTemplateFromFile('HTML');
		
	html.data = JSON.stringify({
		'f1': data[1][2],
		'f2': data[2][2],
		'f3': data[3][2],
		'f4': data[4][2],
		'f5': data[5][2],
		'f6': data[6][2],
		'f13': data[8][2],
		'f9': data[10][2],
		'f10': data[11][2],
		'f12': data[13][2]
	});
	
	var out = html.evaluate()
		.setWidth(500)
		.setHeight(640);
	
	SpreadsheetApp.getUi().showModalDialog(out, 'Set Defaults');
}

function writeDef(value) {
	var ss = SpreadsheetApp.getActiveSpreadsheet(),
		sheet = ss.getSheetByName('Default'),
		form = FormApp.openById('1BaIsWXctQAPlSVMKSennKRc5wz-j3xETu0sZYGf7aoQ'),
		items = form.getItems(FormApp.ItemType.TEXT),
		obj = JSON.parse(value);

	sheet.getRange(2, 3).setValue(obj.f1);
	sheet.getRange(3, 3).setValue(obj.f2);
	sheet.getRange(4, 3).setValue(obj.f3);
	sheet.getRange(5, 3).setValue(obj.f4);
	sheet.getRange(6, 3).setValue(obj.f5);
	sheet.getRange(7, 3).setValue(obj.f6);
	sheet.getRange(9, 3).setValue(obj.f13);
	sheet.getRange(11, 3).setValue(obj.f9);
	sheet.getRange(12, 3).setValue(obj.f10);
	sheet.getRange(14, 3).setValue(obj.f12);

	SpreadsheetApp.flush();

	items[0].asTextItem().setHelpText('Default is "' + obj.f2  + '"');
	items[1].asTextItem().setHelpText('Default is ' +  obj.f3  + '');
	items[2].asTextItem().setHelpText('Default is "' + obj.f4  + '"');
	items[3].asTextItem().setHelpText('Default is "' + obj.f5  + '"');
	items[4].asTextItem().setHelpText('Default is ' +  obj.f6  + '');
	items[5].asTextItem().setHelpText('Default is "' + obj.f13 + '"');
	items[6].asTextItem().setHelpText('Default is "' + obj.f9  + '"');
	items[7].asTextItem().setHelpText('Default is ' +  obj.f10 + '');
	items[9].asTextItem().setHelpText('Default is "' + obj.f12 + '"');
}

function include(filename) {
	return HtmlService.createHtmlOutputFromFile(filename).getContent();
}