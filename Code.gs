var SCRIPT_NAME = 'F1_Email_Generator',
	SCRIPT_VERSION = 'v1.4';

function onOpen() {
	var ui = SpreadsheetApp.getUi();

	ui.createMenu('Mail')
		.addItem('Create HTML', 'showMailPopup')
		.addItem('Set Defaults', 'showFormPopup')
		.addSeparator()
		.addItem('Delete Columns', 'deleteColumns')
		.addToUi();
}

function showMailPopup() {
	var ui = SpreadsheetApp.getUi(),
		ss = SpreadsheetApp.getActive(),
		sheet = ss.getSheetByName('Responses'),
		activeRange = sheet.getActiveRange(),
		values = sheet.getRange(1, activeRange.getColumn(), activeRange.getNumRows()).getValues(),
		mail = HtmlService.createTemplateFromFile('Mail.html'),
		content = {
			'header': {
				'img': {
					'top': String(values[0][0]).trim(),
					'title': String(values[1][0]).trim(),
					'width': String(values[2][0]).trim(),
					'src': String(values[3][0]).trim(),
					'link': String(values[4][0]).trim(),
					'bottom': String(values[5][0]).trim()
				},
				'title': {
					'top': String(values[6][0]).trim(),
					'text': String(values[7][0]).trim(),
					'bottom': String(values[8][0]).trim()
				}
			},
			'body': {
				'heading': {
					'top': String(values[9][0]).trim(),
					'text': String(values[10][0]).trim(),
					'bottom': String(values[11][0]).trim()
				},
				'img': {
					'top': String(values[12][0]).trim(),
					'title': String(values[13][0]).trim(),
					'width': String(values[14][0]).trim(),
					'src': String(values[15][0]).trim(),
					'link': String(values[16][0]).trim(),
					'bottom': String(values[17][0]).trim()
				},
				'subheading': {
					'top': String(values[18][0]).trim(),
					'text': String(values[19][0]).trim(),
					'bottom': String(values[20][0]).trim()
				},
				'box': {
					'top': String(values[21][0]).trim(),
					'text': String(values[22][0]).trim(),
					'bottom': String(values[23][0]).trim()
				}
			},
			'footer': {
				'staff': {
					'top': String(values[24][0]).trim(),
					'workers': [],
					'bottom': String(values[28][0]).trim()
				},
				'unsubscribe': String(values[29][0]).trim()
			}
		};

	for (var i = 25; i <= 27; i++) {
		var name = String(values[i][0]).trim(),
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

	sheet.getRange('D1').setValue(values.header.img.title);
	sheet.getRange('D2').setValue(values.header.img.width);
	sheet.getRange('D3').setValue(values.header.img.src);
	sheet.getRange('D4').setValue(values.header.img.link);

	sheet.getRange('D5').setValue(values.header.title.top);

	sheet.getRange('D7').setValue(values.body.heading.top);
	sheet.getRange('D9').setValue(values.body.heading.bottom);

	sheet.getRange('D10').setValue(values.body.img.title);
	sheet.getRange('D11').setValue(values.body.img.width);
	sheet.getRange('D12').setValue(values.body.img.src);
	sheet.getRange('D13').setValue(values.body.img.link);
	sheet.getRange('D14').setValue(values.body.img.top);
	sheet.getRange('D15').setValue(values.body.img.bottom);
}

function getDefaultValues() {
	var ss = SpreadsheetApp.getActive(),
		sheet = ss.getSheetByName('Defaults'),
		values = sheet.getRange('D1:D30').getValues();

	return {
		'header': {
			'img': {
				'top': String(values[0][0]).trim(),
				'title': String(values[1][0]).trim(),
				'width': String(values[2][0]).trim(),
				'src': String(values[3][0]).trim(),
				'link': String(values[4][0]).trim(),
				'bottom': String(values[5][0]).trim()
			},
			'title': {
				'top': String(values[6][0]).trim(),
				'text': String(values[7][0]).trim(),
				'bottom': String(values[8][0]).trim()
			}
		},
		'body': {
			'heading': {
				'top': String(values[9][0]).trim(),
				'text': String(values[10][0]).trim(),
				'bottom': String(values[11][0]).trim()
			},
			'img': {
				'top': String(values[12][0]).trim(),
				'title': String(values[13][0]).trim(),
				'width': String(values[14][0]).trim(),
				'src': String(values[15][0]).trim(),
				'link': String(values[16][0]).trim(),
				'bottom': String(values[17][0]).trim()
			},
			'subheading': {
				'top': String(values[18][0]).trim(),
				'text': String(values[19][0]).trim(),
				'bottom': String(values[20][0]).trim()
			},
			'box': {
				'top': String(values[21][0]).trim(),
				'text': String(values[22][0]).trim(),
				'bottom': String(values[23][0]).trim()
			}
		},
		'footer': {
			'staff': {
				'top': String(values[24][0]).trim(),
				'workers': [],
				'bottom': String(values[28][0]).trim()
			},
			'unsubscribe': String(values[29][0]).trim()
		}
	};
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