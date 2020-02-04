var SCRIPT_NAME = 'Email_Generator'
var SCRIPT_VERSION = 'v1.8 dev cbd'

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
			'section_1_text': {
				'top': getValue(values, 10),
				'text': getValue(values, 11),
				'bottom': getValue(values, 12)
			},
			'section_1_box': {
				'top': getValue(values,13),
				'text': getValue(values, 14),
				'bottom': getValue(values, 15)
            },
			'section_1_img': {
				'top': getValue(values, 16),
				'title': getValue(values, 17),
				'width': getValue(values, 18),
				'src': getValue(values, 19),
				'link': getValue(values, 20),
				'bottom': getValue(values, 21)
			},
			'section_2_text': {
				'top': getValue(values, 22),
				'text': getValue(values, 23),
				'bottom': getValue(values, 24)
			},
			'section_2_box': {
				'top': getValue(values, 25),
				'text': getValue(values, 26),
				'bottom': getValue(values, 27)
            },
			'section_2_img': {
				'top': getValue(values, 28),
				'title': getValue(values, 29),
				'width': getValue(values, 30),
				'src': getValue(values, 31),
				'link': getValue(values, 32),
				'bottom': getValue(values, 33)
			},
			'section_3_text': {
				'top': getValue(values, 34),
				'text': getValue(values, 35),
				'bottom': getValue(values, 36)
			},
			'section_3_box': {
				'top': getValue(values, 37),
				'text': getValue(values, 38),
				'bottom': getValue(values, 39)
            },
			'section_3_img': {
				'top': getValue(values, 40),
				'title': getValue(values, 41),
				'width': getValue(values, 42),
				'src': getValue(values, 43),
				'link': getValue(values, 44),
				'bottom': getValue(values, 45)
			},
			'section_4_text': {
				'top': getValue(values, 46),
				'text': getValue(values, 47),
				'bottom': getValue(values, 48)
			},
			'section_4_box': {
				'top': getValue(values, 49),
				'text': getValue(values, 50),
				'bottom': getValue(values, 51)
            },
			'section_4_img': {
				'top': getValue(values, 52),
				'title': getValue(values, 53),
				'width': getValue(values, 54),
				'src': getValue(values, 55),
				'link': getValue(values, 56),
				'bottom': getValue(values, 57)
			},

		},
		'footer': {
			'staff': {
				'top': getValue(values, 59),
				'workers': [],
				'bottom': getValue(values, 63)
			},
			'unsubscribe': getValue(values, 64)
		}
	};
}

function generateHtmlEmail() {
	var ui = SpreadsheetApp.getUi(),
		ss = SpreadsheetApp.getActive(),
		sheet = ss.getSheetByName('Input'),
		values = sheet.getRange('F4:F68').getValues(),
		mail = HtmlService.createTemplateFromFile('Mail.html'),
		content = getContentObject(values),
		names = [
			getValue(values, 60),
			getValue(values, 61),
			getValue(values, 62)
		];



	for (var i = 0; i < names.length; i++) {
		var name = names[i],
			nameParts = name.split(' ');

		if (nameParts.length == 2) {
			var person = getStaffObject(nameParts[0], nameParts[1]);

//			if (content.footer.unsubscribe.toUpperCase() != person.team.toUpperCase()) {
//				var resp = ui.alert('Warning', ('According to the Staff Data spreadsheet, ' + person.name + ' is not in ' + content.footer.unsubscribe + '. \n\n Do you wish to continue?'), ui.ButtonSet.YES_NO);
//
//				if (resp == ui.Button.NO) {
//					return;
//				}
//			}

			content.footer.staff.workers.push(person);
		}
	}

	content = mergeObjects(content, getDefaultValues());
  
    content.body.section_1_text['paragraphs'] = content.body.section_1_text.text.split('\n');
	content.body.section_1_box['paragraphs'] = content.body.section_1_box.text.split('\n');

	content.body.section_2_text['paragraphs'] = content.body.section_2_text.text.split('\n');
	content.body.section_2_box['paragraphs'] = content.body.section_2_box.text.split('\n');

	content.body.section_3_text['paragraphs'] = content.body.section_3_text.text.split('\n');
	content.body.section_3_box['paragraphs'] = content.body.section_3_box.text.split('\n');

	content.body.section_4_text['paragraphs'] = content.body.section_4_text.text.split('\n');
	content.body.section_4_box['paragraphs'] = content.body.section_4_box.text.split('\n');

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

/*function deleteColumns() {
	var ss = SpreadsheetApp.getActive(),
		sheet = ss.getSheetByName('Input'),
		start = 4,
		end = sheet.getLastColumn() - (start - 1);

	sheet.deleteColumns(start, end);
}
*/
function setDefaultValues(values) {
	var ss = SpreadsheetApp.getActive(),
		sheet = ss.getSheetByName('Defaults');

	sheet.getRange('F4').setValue(values.header.img.top);
	sheet.getRange('F5').setValue(values.header.img.title);
	sheet.getRange('F6').setValue(values.header.img.width);
	sheet.getRange('F7').setValue(values.header.img.src);
	sheet.getRange('F8').setValue(values.header.img.link);
	sheet.getRange('F9').setValue(values.header.img.bottom);
	sheet.getRange('F10').setValue(values.header.title.top);
	sheet.getRange('F12').setValue(values.header.title.bottom);

	sheet.getRange('F14').setValue(values.body.section_1_text.top);
	sheet.getRange('F16').setValue(values.body.section_1_text.bottom);
  	sheet.getRange('F17').setValue(values.body.section_1_box.top);
	sheet.getRange('F19').setValue(values.body.section_1_box.bottom);
	sheet.getRange('F20').setValue(values.body.section_1_img.top);
	sheet.getRange('F21').setValue(values.body.section_1_img.title);
	sheet.getRange('F22').setValue(values.body.section_1_img.width);
	sheet.getRange('F23').setValue(values.body.section_1_img.src);
	sheet.getRange('F24').setValue(values.body.section_1_img.link);
	sheet.getRange('F25').setValue(values.body.section_1_img.bottom);
  
  	sheet.getRange('F26').setValue(values.body.section_2_text.top);
	sheet.getRange('F28').setValue(values.body.section_2_text.bottom);
  	sheet.getRange('F29').setValue(values.body.section_2_box.top);
	sheet.getRange('F31').setValue(values.body.section_2_box.bottom);
	sheet.getRange('F32').setValue(values.body.section_2_img.top);
	sheet.getRange('F33').setValue(values.body.section_2_img.title);
	sheet.getRange('F34').setValue(values.body.section_2_img.width);
	sheet.getRange('F35').setValue(values.body.section_2_img.src);
	sheet.getRange('F36').setValue(values.body.section_2_img.link);
	sheet.getRange('F37').setValue(values.body.section_2_img.bottom);
    
  	sheet.getRange('F38').setValue(values.body.section_3_text.top);
	sheet.getRange('F40').setValue(values.body.section_3_text.bottom);
  	sheet.getRange('F41').setValue(values.body.section_3_box.top);
	sheet.getRange('F43').setValue(values.body.section_3_box.bottom);
	sheet.getRange('F44').setValue(values.body.section_3_img.top);
	sheet.getRange('F45').setValue(values.body.section_3_img.title);
	sheet.getRange('F46').setValue(values.body.section_3_img.width);
	sheet.getRange('F47').setValue(values.body.section_3_img.src);
	sheet.getRange('F48').setValue(values.body.section_3_img.link);
	sheet.getRange('F49').setValue(values.body.section_3_img.bottom);
  
  	sheet.getRange('F50').setValue(values.body.section_4_text.top);
	sheet.getRange('F52').setValue(values.body.section_4_text.bottom);
  	sheet.getRange('F53').setValue(values.body.section_4_box.top);
	sheet.getRange('F55').setValue(values.body.section_4_box.bottom);
	sheet.getRange('F56').setValue(values.body.section_4_img.top);
	sheet.getRange('F57').setValue(values.body.section_4_img.title);
	sheet.getRange('F58').setValue(values.body.section_4_img.width);
	sheet.getRange('F59').setValue(values.body.section_4_img.src);
	sheet.getRange('F60').setValue(values.body.section_4_img.link);
	sheet.getRange('F61').setValue(values.body.section_4_img.bottom);

	sheet.getRange('F63').setValue(values.footer.staff.top);
	sheet.getRange('F67').setValue(values.footer.staff.bottom);
}


function getDefaultValues() {
	var ss = SpreadsheetApp.getActive(),
		sheet = ss.getSheetByName('Defaults'),
		values = sheet.getRange('F4:F68').getValues(),
		content = getContentObject(values);

	return content;
}

function getStaffObject(firstname, lastname) {
    var staffSheetId = Config.get('STAFF_DATA_GSHEET_ID'),
        ss = SpreadsheetApp.openById(staffSheetId),
		sheet = ss.getSheetByName('Staff Directory'),
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

