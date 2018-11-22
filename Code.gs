var SCRIPT_NAME = 'Email_Generator',
	SCRIPT_VERSION = 'v1.8';

function onOpen() {
	var ui = SpreadsheetApp.getUi();

	ui.createMenu('Mail')
		.addItem('Create HTML', 'showMailPopup')
		.addItem('Set Defaults', 'showFormPopup')
		.addSeparator()
		.addItem('Delete Columns', 'deleteColumns')
		.addToUi();
}

function getValue(values, index) {
	return String(values[index][0]).trim();
}

function getContentObject(values) {
	return {
		'header': {
			'img': {
				'top': getValue(values, 0),
				'title': getValue(values, 1),
				'width': getValue(values, 2),
				'src': getValue(values, 3),
				'link': getValue(values, 4),
				'bottom': getValue(values, 5)
			},
			'title': {
				'top': getValue(values, 6),
				'text': getValue(values, 7),
				'bottom': getValue(values, 8)
			}
		},
		'body': {
			'heading': {
				'top': getValue(values, 10),
				'text': getValue(values, 11),
				'bottom': getValue(values, 12)
			},
			'img': {
				'top': getValue(values, 13),
				'title': getValue(values, 14),
				'width': getValue(values, 15),
				'src': getValue(values, 16),
				'link': getValue(values, 17),
				'bottom': getValue(values, 18)
			},
			'subheading': {
				'top': getValue(values, 19),
				'text': getValue(values, 20),
				'bottom': getValue(values, 21)
			},
			'box': {
				'top': getValue(values, 22),
				'text': getValue(values, 23),
				'bottom': getValue(values, 24)
			}
		},
		'footer': {
			'staff': {
				'top': getValue(values, 26),
				'workers': [],
				'bottom': getValue(values, 30)
			},
			'unsubscribe': getValue(values, 31)
		}
	};
}

function showMailPopup() {
	var ui = SpreadsheetApp.getUi(),
		ss = SpreadsheetApp.getActive(),
		sheet = ss.getSheetByName('Responses'),
		values = sheet.getRange('D3:D34').getValues(),
		mail = HtmlService.createTemplateFromFile('Mail.html'),
		content = getContentObject(values),
		names = [
			getValue(values, 27),
			getValue(values, 28),
			getValue(values, 29)
		];

	for (var i = 0; i < names.length; i++) {
		var name = names[i],
			nameParts = name.split(' ');

		if (nameParts.length == 2) {
			var person = getStaffObject(nameParts[0], nameParts[1]);

			if (content.footer.unsubscribe.toUpperCase() != person.team.toUpperCase()) {
				var resp = ui.alert('Warning', (person.name + ' is not in ' + content.footer.unsubscribe + '. Do you wish to continue?'), ui.ButtonSet.YES_NO);

				if (resp == ui.Button.NO) {
					return;
				}
			}

			content.footer.staff.workers.push(person);
		}
	}

	content = mergeObjects(content, getDefaultValues());

	content.body.heading['paragraphs'] = content.body.heading.text.split('\n');
	content.body.subheading['paragraphs'] = content.body.subheading.text.split('\n');
	content.body.box['paragraphs'] = content.body.box.text.split('\n');

	mail.content = content;

	var html = mail.evaluate()
				   .setWidth(800)
				   .setHeight(640);

	ui.showModalDialog(html, 'Generated mail');
}

function showFormPopup() {
	var ui = SpreadsheetApp.getUi(),
		form = HtmlService.createTemplateFromFile('Form.html');

	form.content = getDefaultValues();

	var html = form.evaluate()
				   .setWidth(520)
				   .setHeight(640);

	ui.showModalDialog(html, 'Set Defaults');
}

function deleteColumns() {
	var ss = SpreadsheetApp.getActive(),
		sheet = ss.getSheetByName('Responses'),
		start = 4,
		end = sheet.getLastColumn() - (start - 1);

	sheet.deleteColumns(start, end);
}

function setDefaultValues(values) {
	var ss = SpreadsheetApp.getActive(),
		sheet = ss.getSheetByName('Defaults');

	sheet.getRange('D3').setValue(values.header.img.top);
	sheet.getRange('D4').setValue(values.header.img.title);
	sheet.getRange('D5').setValue(values.header.img.width);
	sheet.getRange('D6').setValue(values.header.img.src);
	sheet.getRange('D7').setValue(values.header.img.link);
	sheet.getRange('D8').setValue(values.header.img.bottom);

	sheet.getRange('D9').setValue(values.header.title.top);
	sheet.getRange('D11').setValue(values.header.title.bottom);

	sheet.getRange('D13').setValue(values.body.heading.top);
	sheet.getRange('D15').setValue(values.body.heading.bottom);

	sheet.getRange('D16').setValue(values.body.img.top);
	sheet.getRange('D17').setValue(values.body.img.title);
	sheet.getRange('D18').setValue(values.body.img.width);
	sheet.getRange('D19').setValue(values.body.img.src);
	sheet.getRange('D20').setValue(values.body.img.link);
	sheet.getRange('D21').setValue(values.body.img.bottom);

	sheet.getRange('D22').setValue(values.body.subheading.top);
	sheet.getRange('D24').setValue(values.body.subheading.bottom);

	sheet.getRange('D25').setValue(values.body.box.top);
	sheet.getRange('D27').setValue(values.body.box.bottom);

	sheet.getRange('D29').setValue(values.footer.staff.top);
	sheet.getRange('D33').setValue(values.footer.staff.bottom);
}

function getDefaultValues() {
	var ss = SpreadsheetApp.getActive(),
		sheet = ss.getSheetByName('Defaults'),
		values = sheet.getRange('D3:D34').getValues(),
		content = getContentObject(values);

	return content;
}

function getStaffObject(firstname, lastname) {
	var ss = SpreadsheetApp.openById('1iiFmdqUd-CoWtUjZxVgGcNb74dPVh-l5kuU_G5mmiHI'),
		sheet = ss.getSheetByName('Staff'),
		values = sheet.getDataRange().getValues(),
		person = {
			'name': firstname + ' ' + lastname,
			'title': '',
			'team': '',
			'photo': getStaffImage(firstname, lastname)
		};

	for (var i = 2; i < values.length; i++) {
		if ((values[i][0].toUpperCase() == firstname.toUpperCase()) && (values[i][1].toUpperCase() == lastname.toUpperCase())) {
			person.title = values[i][4];
			person.team = values[i][11];

			break;
		}
	}

	return person;
}

function getStaffImage(firstname, lastname) {
	var folders = DriveApp.getFoldersByName(lastname + ', ' + firstname),
		imgFile = searchFileInFolder(folders, 'BubbleHead');

	if (imgFile != null) {
		var fileId = imgFile.getId();

		return ('https://drive.google.com/uc?export=view&id=' + fileId);
	}

	return 'https://vignette.wikia.nocookie.net/citrus/images/6/60/No_Image_Available.png';
}

function searchFileInFolder(folders, filename) {
	while (folders.hasNext()) {
		var folder = folders.next(),
			files = folder.searchFiles('title contains "' + filename + '"');

		if (files.hasNext()) {
			return files.next();
		}

		var subfolders = folder.getFolders();

		if (subfolders.hasNext()) {
			var file = searchFileInFolder(subfolders, filename);

			if (file != null) {
				return file;
			}
		}
	}

	return;
}

function mergeObjects(obj, src) {
	for (var key in obj) {
		if (obj[key].constructor == Object) {
			obj[key] = mergeObjects(obj[key], src[key]);
		} else if (String(obj[key]) == '') {
			obj[key] = src[key];
		}
	}

	return obj;
}