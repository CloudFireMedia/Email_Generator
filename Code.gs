var SCRIPT_NAME = 'F1_Email_Generator.gdscript',
	SCRIPT_VERSION = 'v0.dev_msk';

function getDefault() {
	var ss = SpreadsheetApp.getActiveSpreadsheet(),
		sheet = ss.getSheetByName('Default'),
		data = sheet.getDataRange().getValues();

	return {
		'filename': data[1][2],
		'HIA': data[2][2],
		'HIW': data[3][2],
		'HIS': data[4][2],
		'HIL': data[5][2],
		'HIP': data[6][2],
		'BTP': data[8][2],
		'BIA': data[10][2],
		'BIW': data[11][2],
		'BIL': data[13][2]
	};
}

function onFormSubmit(e) {
	var objD = getDefault(),
		htmlBody = HtmlService.createHtmlOutputFromFile('HTML1').getContent(),
		htmlTemp;

	if (e.values[2] != '') {
		objD.HIW = String(e.values[2]).replace(/[^0-9]/g, '');
	}

	var HIA = 'Header Image';
	if (e.values[1] != '') {
		HIA = e.values[1];
	}

	if (e.values[3] != '') {
		objD.HIS = e.values[3];
	}

	if (e.values[4] != '') {
		objD.HIL = e.values[4];
	}

	htmlTemp = HtmlService.createHtmlOutputFromFile('HeaderImage').getContent()
		.replace('{HEADER IMAGE WIDTH}', objD.HIW)
		.replace('{HEADER IMAGE URL}', objD.HIS)
		.replace('{HEADER IMAGE LINK URL}', objD.HIL)
		.replace('{HEADER IMAGE ALTERNATE TEXT}', HIA)
		.replace('{HEADER IMAGE ALTERNATE TEXT}', HIA);
	htmlBody = htmlBody.replace('{BEGIN HEADER IMAGE}', htmlTemp);
	htmlTemp = '';

	if (e.values[6] != '') {
		objD.filename = e.values[6];

		if (e.values[5] != '') {
			objD.HIP = String(e.values[5]).replace(/[^0-9]/g, '');
		}

		htmlTemp = HtmlService.createHtmlOutputFromFile('HeaderText').getContent()
			.replace('{HEADER TEXT TOP PADDING}', objD.HIP)
			.replace('{HEADER TEXT}', e.values[6]);
	}

	htmlBody = htmlBody.replace('{BEGIN HEADER TEXT}', htmlTemp);
	htmlTemp = '';

	if (e.values[8] != '') {
		var tSt = setParagraphs(String(e.values[8]));

		htmlTemp = tSt;

		if (e.values[7] != '') {
			htmlTemp = htmlTemp.replace('padding-top: 5px;', 'padding-top: ' + e.values[7] + 'px;')
		} else if (objD.BTP != '') {
			htmlTemp = htmlTemp.replace('padding-top: 5px;', 'padding-top: ' + objD.BTP + 'px;')
		}
	}

	htmlBody = htmlBody.replace('{BEGIN BODY PARAGRAPH #1}', htmlTemp);
	htmlTemp = '';

	if (e.values[11] != '') {
		if (e.values[10] != '') {
			objD.BIW = String(e.values[10]).replace(/[^0-9]/g, '');
		}

		if (e.values[9] != '') {
			objD.BIA = e.values[9];
		}

		if (e.values[12] != '') {
			objD.BIL = e.values[12];
		}

		htmlTemp = HtmlService.createHtmlOutputFromFile('BodyImage').getContent()
			.replace('{BODY IMAGE URL}', e.values[11])
			.replace('{BODY IMAGE ALTERNATE TEXT}', objD.BIA)
			.replace('{BODY IMAGE ALTERNATE TEXT}', objD.BIA)
			.replace('{BODY IMAGE LINK URL}', objD.BIL)
			.replace('{BODY IMAGE WIDTH}', objD.BIW);
	}

	htmlBody = htmlBody.replace('{BEGIN BODY IMAGE}', htmlTemp);
	htmlTemp = '';

	if (e.values[13] != '') {
		var tSt = setParagraphs(String(e.values[13]));

		htmlTemp = tSt;
	}

	htmlBody = htmlBody.replace('{BEGIN BODY PARAGRAPH #4}', htmlTemp);
	htmlTemp = '';

	if (e.values[15] != '') {
		htmlTemp = createStaffHTML(1, e.values[15]);
	}

	htmlBody = htmlBody.replace('{BEGIN STAFF SIGNATURE #1}', htmlTemp);
	htmlTemp = '';

	if (e.values[16] != '') {
		htmlTemp = createStaffHTML(2, e.values[16]);
	}

	htmlBody = htmlBody.replace('{BEGIN STAFF SIGNATURE #2}', htmlTemp);
	htmlTemp = '';

	if (e.values[17] != '') {
		htmlTemp = createStaffHTML(3, e.values[17]);
	}

	htmlBody = htmlBody.replace('{BEGIN STAFF SIGNATURE #3}', htmlTemp)
					   .replace('{Email List}', e.values[18])
					   .replace('{Email List}', e.values[18]);
	htmlTemp = '';

	if (!matchDates(e.values[8] + ' ' + e.values[13])) {
		htmlTemp = HtmlService.createHtmlOutputFromFile('NoDate').getContent();
	}

	//Gray box
	htmlBody = htmlBody.replace('{line1}', htmlTemp);
	htmlTemp = '';

	if (e.values[14] != '') {
		htmlTemp = HtmlService.createHtmlOutputFromFile('GrayBox').getContent()
			.replace('{Gray Box Text}', e.values[14]);
	}

	htmlBody = htmlBody.replace('{BEGIN GBOX} ', htmlTemp);

	try {
		var ui = SpreadsheetApp.getUi()
	} catch (e) {
		//
	}

	if (ui) {
		var html = HtmlService.createTemplate('<xmp>' + htmlBody + '</xmp><br><hr><br>' + htmlBody),
			out = html.evaluate()
					  .setWidth(800)
					  .setHeight(640);

		ui.showModalDialog(out, 'HTML code');
	} else {
		var doc = DocumentApp.create(objD.filename);

		doc.getBody().setText(htmlBody);
		doc.saveAndClose();
	}
}

function createStaffHTML(num, value) {
	var obj = getStaffSignature(String(value)),
		result = HtmlService.createHtmlOutputFromFile('Signature' + num).getContent()
			.replace('{BUBBLE HEAD URL FOR STAFF #' + num + '}', obj.img)
			.replace('{NAME OF STAFF #' + num + '}', value)
			.replace('{NAME OF STAFF #' + num + '}', value)
			.replace('{NAME OF STAFF #' + num + '}', value)
			.replace('{JOB TITLE FOR STAFF #' + num + '}', obj.title);

	return result;
}

function getStaffSignature(value) {
	var ss = SpreadsheetApp.openById('1iiFmdqUd-CoWtUjZxVgGcNb74dPVh-l5kuU_G5mmiHI'),
		sheet = ss.getSheetByName('Staff'),
		data = sheet.getDataRange().getValues(),
		value = value.trim().toUpperCase(),
		result = {
			'title': '',
			'alt': '',
			'img': ''
		};

loop:
	for (var i = 2; i < data.length; i++) {
		var aaa = (data[i][0] + ' ' + data[i][1]).trim().toUpperCase();

		if (aaa == value) {
			var folderName = data[i][1] + ', ' + data[i][0],
				folders = DriveApp.getFoldersByName(folderName),
				filename = 'BubbleHead';

			result['title'] = data[i][4];
			result['alt'] = data[i][0] + ' ' + data[i][1];

			if (folders.hasNext()) {
				var folder = folders.next(),
					foldersI = folder.getFolders();

				while (foldersI.hasNext()) {
					var folderI = foldersI.next(),
						files = folderI.searchFiles('title contains "' + filename + '"');

					if (files.hasNext()) {
						var file = files.next();

						result['img'] = 'https://drive.google.com/uc?export=view&id=' + file.getId();

						break loop;
					}
				}
			}

			break;
		}
	}

	return null;
}

function setParagraphs(value) {
	if (value == '') {
		return '';
	}

	var htmlTemp = HtmlService.createHtmlOutputFromFile('BodyParagraph1').getContent(),
		value = processNew(value);

	if (value.indexOf('\n') >= 0) {
		var paragraphs = value.split('\n'),
			result = '';

		for (var i = 0; i < paragraphs.length; i++) {
			var aaa = String(paragraphs[i]).trim();

			if ((aaa != '') && (aaa != '<span style="background-color: #ffff00; font-weight: bold;"></span>')) {
				result += htmlTemp.replace('{BODY PARAGRAPH #1}', aaa);
			}
		}

		return result;
	}

	return htmlTemp.replace('{BODY PARAGRAPH #1}', value.trim());
}

function createLinks(value) {
	var str = String(value),
		reg = new RegExp('(https?:\\/\\/[^\\s]+)\\s\\((.*)\\)', 'i'),
		res = reg.exec(str);

	while (res != null && res.length >= 3) {
		str = str.replace(reg, ('<a href="' + res[1] + '">' + res[2] + '</a>'));
		res = reg.exec(str);
	}

	return str;
}

//OLD
function formatParagrahs(strString) {
	var strRet = '';

	if (('' + strString) == '') {
		//
	} else if (('' + strString).indexOf('\n') >= 0) {
		var arr = ('' + strString).split('\n');

		for (var i = 0; i < arr.length; i++) {
			if (('' + arr[i]).trim() != '') {
				var htmlTemp = HtmlService.createHtmlOutputFromFile('BodyParagraph1').getContent();

				htmlTemp = htmlTemp.replace('{BODY PARAGRAPH #1}', ('' + arr[i]).trim());
				strRet = strRet + htmlTemp
			}
		}
	} else {
		var htmlTemp = HtmlService.createHtmlOutputFromFile('BodyParagraph1').getContent();

		htmlTemp = htmlTemp.replace('{BODY PARAGRAPH #1}', ('' + strString).trim());
		strRet = htmlTemp
	}

	return strRet;
}

function matchDates(str) {
	var mNames = [
			'January',
			'February',
			'March',
			'April',
			'May',
			'June',
			'July',
			'August',
			'September',
			'October',
			'November',
			'December',
			'September'
		],
		mAbr = [
			'Jan',
			'Feb',
			'Mar',
			'Apr',
			'May',
			'Jun',
			'Jul',
			'Aug',
			'Sep',
			'Oct',
			'Nov',
			'Dec',
			'Sept'
		];

	for (var i = 0; i < mAbr.length; i++) {
		var regexp = new RegExp('(' + mAbr[i] + '|' + mNames[i] + ')\\s[0-3]?[0-9](\\sand\\s[0-3]?[0-9])?(,)?(\\s[1-2][0-9][0-9][0-9])?\\sat\\s[0-2]?[0-9](:)?([0-5][0-9])?(a|p)m', "ig");

		if (regexp.test(str) == true) {
			return true;
		}
	}

	return false;
}

function processNew(inistr) {
	var reg0 = new RegExp('\\[\\!\\!([^\\]]+)\\!\\!\\]\\s?\\(([^\\)]+)\\)', 'igm'),
		reg = new RegExp('\\[([^\\]]+)\\]\\s?\\(([^\\)]+)\\)', 'igm'),
		reg1 = new RegExp('\\!\\!([^(?:\\!\\!)]+)\\!\\!', 'igm'),
		myArray = [],
		outtext = inistr;

	while ((myArray = reg0.exec(inistr)) !== null) {
		var str2Replace = myArray[0],
			str2KeepT = ('' + myArray[1]).trim().replace(/\n+/g, ''),
			str2KeepL = ('' + myArray[2]).trim(),
			str2KeepLl = str2KeepL.toLowerCase();

		if (str2KeepLl.indexOf('http') < 0) {
			str2KeepL = 'http://' + str2KeepL;
		}

		var strLink = '<a href="' + str2KeepL + '" style="background-color: #ffff00; font-weight: bold;" >' + str2KeepT + '</a>';
		
		outtext = outtext.replace(str2Replace, strLink)
	}

	inistr = outtext;

	while ((myArray = reg.exec(inistr)) !== null) {
		var str2Replace = myArray[0],
			str2KeepT = String(myArray[1]).trim().replace(/\n+/g, ''),
			str2KeepL = String(myArray[2]).trim(),
			str2KeepLl = str2KeepL.toLowerCase();

		if (str2KeepLl.indexOf('http') < 0) {
			str2KeepL = 'http://' + str2KeepL;
		}

		var strLink = '<a href="' + str2KeepL + '" >' + str2KeepT + '</a>';

		outtext = outtext.replace(str2Replace, strLink);
	}

	outtext = outtext.replace(/\n+/g, '\n');
	inistr = outtext;

	while ((myArray = reg1.exec(inistr)) !== null) {
		var str2Replace = myArray[0],
			str2Keep = String(myArray[1]).trim();

		str2Keep = String(str2Keep).replace(/\n/g, '</span>\n<span style="background-color: #ffff00; font-weight: bold;">');

		var strSpan = '<span style="background-color: #ffff00; font-weight: bold;">' + str2Keep + '</span>';

		outtext = outtext.replace(str2Replace, strSpan);
	}

	return outtext;
}