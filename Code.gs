var SCRIPT_NAME = 'F1_Email_Generator.gdscript'
var SCRIPT_VERSION = 'v0.dev_msk'

function getDefault() {

	var sp = SpreadsheetApp.getActiveSpreadsheet();
	var sh = sp.getSheetByName('Default');
	var data = sh.getDataRange().getValues();

	var obj = {
		filename: data[1][2],
		HIA: data[2][2],
		HIW: data[3][2],
		HIS: data[4][2],
		HIL: data[5][2],
		HIP: data[6][2],
		BTP: data[8][2],
		BIA: data[10][2],
		BIW: data[11][2],
		BIL: data[13][2],
	};

	return obj;

}

function onFormSubmit(e) {

	var objD = getDefault()

		//var filename = 'F1 Email.txt'
		//Logger.log(JSON.stringify(objD))
		//Logger.log(objD.filename)
		//Logger.log(e.values)
		//Logger.log(e.namedValues)
		var htmlBody = HtmlService.createHtmlOutputFromFile('HTML1').getContent();

	var htmlTemp = "";
	//added by Chad: the script should NOT remove this section if [3] does not contain content
	// if (e.values[3] != "")
	// {

	//var HIW = 560; //default
	if (e.values[2] != "") {
		objD.HIW = ('' + e.values[2]).replace(/[^0-9]/g, '');
	}

	var HIA = "Header Image"; //default
	if (e.values[1] != "") {
		HIA = e.values[1];
	}

	//added by Chad: this should be the default Header Image Source.
	//var HIS = "http://www.christchurchnashville.org/wp-content/uploads/2017/10/Header-Logo-1.png"; //default
	if (e.values[3] != "") {
		objD.HIS = e.values[3];
	}

	//var HIL = "#"; //default
	if (e.values[4] != "") {
		objD.HIL = e.values[4];
	}

	var htmlTemp = HtmlService.createHtmlOutputFromFile('HeaderImage').getContent();
	htmlTemp = htmlTemp.replace("{HEADER IMAGE WIDTH}", objD.HIW).replace("{HEADER IMAGE URL}", objD.HIS)
		.replace("{HEADER IMAGE LINK URL}", objD.HIL).replace("{HEADER IMAGE ALTERNATE TEXT}", HIA).replace("{HEADER IMAGE ALTERNATE TEXT}", HIA);

	//	}

	htmlBody = htmlBody.replace("{BEGIN HEADER IMAGE}", htmlTemp);

	var htmlTemp = "";

	if (e.values[6] != "") {
		//Logger.log(e.values[6])
		//Logger.log(typeof e.values[6])
		objD.filename = e.values[6];

		//var HIP = 25; //default
		if (e.values[5] != "") {
			objD.HIP = ('' + e.values[5]).replace(/[^0-9]/g, '');
		}

		var htmlTemp = HtmlService.createHtmlOutputFromFile('HeaderText').getContent();
		htmlTemp = htmlTemp.replace("{HEADER TEXT TOP PADDING}", objD.HIP).replace("{HEADER TEXT}", e.values[6]);

	}

	htmlBody = htmlBody.replace("{BEGIN HEADER TEXT}", htmlTemp);

	var htmlTemp = "";

	if (e.values[8] != "") {

		var tSt = setParagraphs(e.values[8]);
		htmlTemp = tSt;

		//Logger.log(e.values[7])
		//Logger.log(objD.BTP)
		if (e.values[7] != "") {

			htmlTemp = htmlTemp.replace('padding-top: 5px;', 'padding-top: ' + e.values[7] + 'px;')

		} else if (objD.BTP != "") {

			htmlTemp = htmlTemp.replace('padding-top: 5px;', 'padding-top: ' + objD.BTP + 'px;')
		}

	}

	//Logger.log(htmlTemp)

	htmlBody = htmlBody.replace("{BEGIN BODY PARAGRAPH #1}", htmlTemp);

	var htmlTemp = "";

	if (e.values[11] != "") {

		//var BIW = 560; //default
		if (e.values[10] != "") {
			objD.BIW = ('' + e.values[10]).replace(/[^0-9]/g, '');
		}

		//var BIA = "Body Image"; //default
		if (e.values[9] != "") {
			objD.BIA = e.values[9];
		}

		//var BIL = "#"; //default
		if (e.values[12] != "") {
			objD.BIL = e.values[12];
		}

		var htmlTemp = HtmlService.createHtmlOutputFromFile('BodyImage').getContent();
		htmlTemp = htmlTemp.replace("{BODY IMAGE URL}", e.values[11]).replace("{BODY IMAGE ALTERNATE TEXT}", objD.BIA).replace("{BODY IMAGE ALTERNATE TEXT}", objD.BIA)
			.replace("{BODY IMAGE LINK URL}", objD.BIL).replace("{BODY IMAGE WIDTH}", objD.BIW);

	}

	htmlBody = htmlBody.replace("{BEGIN BODY IMAGE}", htmlTemp);

	var htmlTemp = "";

	if (e.values[13] != "") {

		var tSt = setParagraphs(e.values[13]);
		htmlTemp = tSt;

	}

	htmlBody = htmlBody.replace("{BEGIN BODY PARAGRAPH #4}", htmlTemp);

	var htmlTemp = "";

	if (e.values[15] != "") {

		var obj = getStaffSignature(e.values[15]);

		var htmlTemp = HtmlService.createHtmlOutputFromFile('Signature1').getContent();
		htmlTemp = htmlTemp.replace("{BUBBLE HEAD URL FOR STAFF #1}", obj.img).replace("{NAME OF STAFF #1}", e.values[15])
			.replace("{NAME OF STAFF #1}", e.values[15]).replace("{NAME OF STAFF #1}", e.values[15]).replace("{JOB TITLE FOR STAFF #1}", obj.title);

	}

	htmlBody = htmlBody.replace(" {BEGIN STAFF SIGNATURE #1}", htmlTemp);

	var htmlTemp = "";

	if (e.values[16] != "") {

		var obj = getStaffSignature(e.values[16]);

		var htmlTemp = HtmlService.createHtmlOutputFromFile('Signature2').getContent();
		htmlTemp = htmlTemp.replace("{BUBBLE HEAD URL FOR STAFF #2}", obj.img).replace("{NAME OF STAFF #2}", e.values[16])
			.replace("{NAME OF STAFF #2}", e.values[16]).replace("{NAME OF STAFF #2}", e.values[16]).replace("{JOB TITLE FOR STAFF #2}", obj.title);

	}

	htmlBody = htmlBody.replace(" {BEGIN STAFF SIGNATURE #2}", htmlTemp);

	var htmlTemp = "";

	if (e.values[17] != "") {

		var obj = getStaffSignature(e.values[17]);

		var htmlTemp = HtmlService.createHtmlOutputFromFile('Signature3').getContent();
		htmlTemp = htmlTemp.replace("{BUBBLE HEAD URL FOR STAFF #3}", obj.img).replace("{NAME OF STAFF #3}", e.values[17])
			.replace("{NAME OF STAFF #3}", e.values[17]).replace("{NAME OF STAFF #3}", e.values[17]).replace("{JOB TITLE FOR STAFF #3}", obj.title);

	}

	htmlBody = htmlBody.replace(" {BEGIN STAFF SIGNATURE #3}", htmlTemp);

	htmlBody = htmlBody.replace("{Email List}", e.values[18]);
	htmlBody = htmlBody.replace("{Email List}", e.values[18]);

	var htmlTemp = "";

	var str = e.values[8] + ' ' + e.values[13];

	if (matchDates(str) != true) {
		htmlTemp = HtmlService.createHtmlOutputFromFile('NoDate').getContent();
	}

	htmlBody = htmlBody.replace("{line1}", htmlTemp);

	//Gray box

	var htmlTemp = "";

	if (e.values[14] != "") {

		var strGB = e.values[14];

		var htmlTemp = HtmlService.createHtmlOutputFromFile('GrayBox').getContent();
		htmlTemp = htmlTemp.replace("{Gray Box Text}", strGB);

	}
	htmlBody = htmlBody.replace("{BEGIN GBOX} ", htmlTemp);

	/*
	var strEmailL='subject=Unsubscribe me from '+ e.values[16] +' Emails';
	htmlBody = htmlBody.replace("{Email List}", encodeURIComponent(strEmailL));

	


	var email = 'ioana.prof@gmail.com';
	var subject = 'Test subject';
	var body = 'Test body';


	//Logger.log(htmlBody)

	MailApp.sendEmail(email, subject, body, {
	'htmlBody' : htmlBody
	})

 */

	try {
		var ui = SpreadsheetApp.getUi()

	} catch (e) {}

	if (ui) {
		var html = HtmlService.createTemplate('<xmp>' + htmlBody + '</xmp><br><hr><br>' + htmlBody)

			var out = html.evaluate()
			.setWidth(800)
			.setHeight(640);

		ui.showModalDialog(out, 'HTML code');

	} else {

		//Logger.log(objD.filename)
		var doc = DocumentApp.create(objD.filename);
		doc.getBody().setText(htmlBody);
		doc.saveAndClose();
	}
}

function getStaffSignature(strStr) {

	var response = {
		'title': '',
		'alt': '',
		'img': '',
	}

	var strIni = '' + strStr;

	var sp = SpreadsheetApp.openById('1iiFmdqUd-CoWtUjZxVgGcNb74dPVh-l5kuU_G5mmiHI');
	var sh = sp.getSheetByName('Staff');
	var data = sh.getDataRange().getValues();

	loop:
	for (var i = 2; i < data.length; i++) {

		if (('' + data[i][0] + ' ' + data[i][1]).trim().toUpperCase() == strIni.trim().toUpperCase()) {

			response.title = data[i][4];
			response.alt = data[i][0] + " " + data[i][1];

			var folderName = '' + data[i][1] + ', ' + data[i][0];
			var filename = 'BubbleHead';

			var folders = DriveApp.getFoldersByName(folderName);
			if (folders.hasNext()) {
				var folder = folders.next();
				// Logger.log("folder found");

				var foldersI = folder.getFolders();
				while (foldersI.hasNext()) {
					var folderI = foldersI.next();

					var files = folderI.searchFiles('title contains "' + filename + '"');
					if (files.hasNext()) {
						var file = files.next();
						response.img = "https://drive.google.com/uc?export=view&id=" + file.getId();
						break loop;
					}

				}

			}

			break;
		}
	}

	return response;

}

function setParagraphs(strString) {

	var strRet = '' + strString;

	if (strRet == '') {
		return '';
	}

	/**
	var mark = "!!";
	var strDiv = '<span style="background-color: #ffff00; font-weight: bold;">';

	// check bold
	var count = (strRet.match(/!!/g) || []).length;

	//Logger.log(count)
	if (count > 0) {
	//Logger.log(count)
	//complete tags
	for (var i = 0; i < count; i = i + 2) {
	//Logger.log(i)
	strRet = strRet.replace('!!', strDiv)
	if ((i + 1) < count) {
	strRet = strRet.replace('!!', '</span>')
	}

	}

	if (Math.abs(count % 2) == 1) {
	strRet = strRet + '</span>'

	}
	//split by pharagraphs
	if (strRet.indexOf('\n') >= 0) {
	var start = strRet.length;
	//Logger.log(strRet);
	//Logger.log("stage2");
	while (start >= 0) {

	var index1 = strRet.lastIndexOf('</span>', start);
	if (index1 < 0) {
	start = index1
	} else {
	var index2 = strRet.lastIndexOf(strDiv, index1);
	if (index2 < 0) {
	start = index2
	} else {
	var indexS = strRet.lastIndexOf('\n', index1);
	while (indexS >= 0) {
	//Logger.log(index1)
	//Logger.log(indexS)
	//Logger.log(index2)

	if (index1 > indexS && index2 < indexS) {
	strRet = strRet.substring(0, indexS) + '</span>\n' + strDiv + strRet.substring(1 + indexS);
	//Logger.log(indexS)
	//Logger.log("Success")
	}
	strRet

	indexS = strRet.lastIndexOf('\n', indexS - 1);
	}
	start = index2;
	}
	}
	}
	}
	}
	 **/
	//Logger.log(strRet)
	var strString = processNew(strRet);
	var strResp = '';

	// split by pharagraphs
	if (('' + strString).indexOf('\n') >= 0) {

		var arr = ('' + strString).split('\n');
		//Logger.log(arr)
		for (var i = 0; i < arr.length; i++) {
			if (('' + arr[i]).trim() != '' && ('' + arr[i]).trim() != ('<span style="background-color: #ffff00; font-weight: bold;"></span>')) {
				var htmlTemp = HtmlService.createHtmlOutputFromFile('BodyParagraph1').getContent();
				htmlTemp = htmlTemp.replace("{BODY PARAGRAPH #1}", ('' + arr[i]).trim());
				strResp = strResp + htmlTemp;
			}

		}

	} else {

		var htmlTemp = HtmlService.createHtmlOutputFromFile('BodyParagraph1').getContent();
		htmlTemp = htmlTemp.replace("{BODY PARAGRAPH #1}", ('' + strString).trim());
		strResp = htmlTemp;
	}

	return strResp;

}

function createLinks(strStr) {

	var str = '' + strStr;
	var reg = new RegExp('(https?:\\/\\/[^\\s]+)\\s\\((.*)\\)', "i");

	var res = reg.exec(str);
	//Logger.log(res)
	while (res != null && res.length >= 3) {

		var link = '<a href="' + res[1] + '">' + res[2] + '</a>';
		str = str.replace(reg, link);
		// Logger.log(str)
		res = reg.exec(str);
		// Logger.log(res)

	}

	return (str);

}

//OLD

function formatParagrahs(strString) {

	var strRet = '';

	if (('' + strString) == '') {}
	else if (('' + strString).indexOf('\n') >= 0) {

		var arr = ('' + strString).split('\n');
		//Logger.log(arr)
		for (var i = 0; i < arr.length; i++) {
			if (('' + arr[i]).trim() != '') {
				var htmlTemp = HtmlService.createHtmlOutputFromFile('BodyParagraph1').getContent();
				htmlTemp = htmlTemp.replace("{BODY PARAGRAPH #1}", ('' + arr[i]).trim());
				strRet = strRet + htmlTemp
			}

		}

	} else {

		var htmlTemp = HtmlService.createHtmlOutputFromFile('BodyParagraph1').getContent();
		htmlTemp = htmlTemp.replace("{BODY PARAGRAPH #1}", ('' + strString).trim());
		strRet = htmlTemp
	}

	return strRet

}


function matchDates(str) {

	//var str="Jan 31 at 12am "
	//var str="January 1 at 1pm "
	//var str="January 1 at 1:00pm"

	//var str="Monday, Jan 1 at 1pm"
	//var str="Monday, January 1 at 1pm"
	//var str="Monday, January 1 at 1:00pm"

	//var str="Monday, Jan 1 2018 at 1pm"
	//var str="Monday, January 1 2018 at 1pm"
	//var str="Monday, January 1 2018 at 1:00pm"

	//var str="Monday, Jan 1, 2018 at 1pm"
	//var str="Monday, January 1, 2018 at 1pm "
	//var str="Monday, January 1, 2018 at 1:00pm"
	//var str="Monday, December 31, 2000 at 23:59pm"
	//var str="December 22 and 23 at 7pm"
	//var dName=['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'];


	var mNames = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December', 'September'];
	var mAbr = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', 'Sept'];

	for (var i = 0; i < mAbr.length; i++) {
		var regexp = new RegExp('(' + mAbr[i] + '|' + mNames[i] + ')\\s[0-3]?[0-9](\\sand\\s[0-3]?[0-9])?(,)?(\\s[1-2][0-9][0-9][0-9])?\\sat\\s[0-2]?[0-9](:)?([0-5][0-9])?(a|p)m', "ig");

		if (regexp.test(str) == true) {
			//Logger.log(true);
			return true;
		}
	}

	return false;

}

function processNew(inistr) {

	//var inistr = "classics and brand [!!this text will be bolded, highlighted, and linked!!](www.website.com) new arrangements. We are especially excited to announce acclaimed Southern Gospel trio, [this text will be linked](http://www.martinsonline.com) (The Martins), as our special guest. \n" +
	//	"We look forward to sharing this wonderful evening with you Saturday and Sunday, December 2 and 3 at 7pm! Tickets are available !!this text will be bolded \n and highlighted!! online http://www.ccnash.org (here) or by calling our church"

	var reg0 = new RegExp('\\[\\!\\!([^\\]]+)\\!\\!\\]\\s?\\(([^\\)]+)\\)', "igm");
	var reg = new RegExp('\\[([^\\]]+)\\]\\s?\\(([^\\)]+)\\)', "igm");
	var reg1 = new RegExp('\\!\\!([^(?:\\!\\!)]+)\\!\\!', "igm");
	// var res = reg.exec(inistr);

	var myArray;

	var outtext = inistr;
	//Logger.log("0\n");
	while ((myArray = reg0.exec(inistr)) !== null) {

		var str2Replace = myArray[0]
			var str2KeepT = ('' + myArray[1]).trim().replace(/\n+/g, '');
		var str2KeepL = ('' + myArray[2]).trim();
		var str2KeepLl = str2KeepL.toLowerCase();
		if (str2KeepLl.indexOf("http") < 0) {
			str2KeepL = "http://" + str2KeepL;
		}

		var strLink = '<a href="' + str2KeepL + '" style="background-color: #ffff00; font-weight: bold;" >' + str2KeepT + '</a>';
		outtext = outtext.replace(str2Replace, strLink)
			//var msg = 'Found ' + myArray[0] + ' \n\n ' + myArray[1] + '\n\n' + myArray[2] + '\n\n';
			//msg += 'Next match starts at ' + reg0.lastIndex;
			//Logger.log(msg);
	}

	inistr = outtext;
	//Logger.log("1\n");
	while ((myArray = reg.exec(inistr)) !== null) {
		var str2Replace = myArray[0]
			var str2KeepT = ('' + myArray[1]).trim().replace(/\n+/g, '');
		var str2KeepL = ('' + myArray[2]).trim();
		var str2KeepLl = str2KeepL.toLowerCase();
		if (str2KeepLl.indexOf("http") < 0) {
			str2KeepL = "http://" + str2KeepL;
		}

		var strLink = '<a href="' + str2KeepL + '" >' + str2KeepT + '</a>';
		outtext = outtext.replace(str2Replace, strLink)
			//var msg = 'Found ' + myArray[0] + ' \n\n ' + myArray[1] + '\n\n' + myArray[2] + '\n\n';
			//msg += 'Next match starts at ' + reg.lastIndex;
			//Logger.log(msg);
	}

	//Logger.log("2\n");

	outtext = outtext.replace(/\n+/g, '\n');
	inistr = outtext;
	while ((myArray = reg1.exec(inistr)) !== null) {
		var str2Replace = myArray[0]
			var str2Keep = ('' + myArray[1]).trim();
		str2Keep = ('' + str2Keep).replace(/\n/g, '</span>\n<span style="background-color: #ffff00; font-weight: bold;">');
		var strSpan = '<span style="background-color: #ffff00; font-weight: bold;">' + str2Keep + '</span>';
		outtext = outtext.replace(str2Replace, strSpan)

			//var msg = '\n  Found \n ' + myArray[0] + ' \n' + myArray[1] + '\n\n';
			//Logger.log(msg);
	}

	//Logger.log(outtext);

	return outtext;
}

function debugNew1() {
	var sp = SpreadsheetApp.getActiveSpreadsheet();
	Logger.log(sp.getId())
}

function debug() {

	Logger.log(getStaffSignature("Chad Barlow"));

}

function debugP() {

	Logger.log(formatParagrahs("Chad Barlow\nChadBarlow"));

}

function debugB() {

	var inistr = "classics and brand new arrangements. We are especially excited to announce acclaimed Southern Gospel trio, http://www.martinsonline.com (The Martins), as our special guest. \n" +
		"!!We look forward to sharing this wonderful evening with you Saturday and Sunday, December 2 and 3 at 7pm!!! Tickets are available online http://www.ccnash.org (here) or by calling our church"

		Logger.log(createLinks(inistr));

}