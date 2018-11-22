var SCRIPT_NAME = 'Email_Generator';
var SCRIPT_VERSION = 'v1.9';

function getValue_(values, index) {
  return String(values[index][0]).trim();
}

function getContentObject_(values) {
  return {
    'header': {
      'img': {
        'top': getValue_(values, 0),
        'title': getValue_(values, 1),
        'width': getValue_(values, 2),
        'src': getValue_(values, 3),
        'link': getValue_(values, 4),
        'bottom': getValue_(values, 5)
      },
      'title': {
        'top': getValue_(values, 6),
        'text': getValue_(values, 7),
        'bottom': getValue_(values, 8)
      }
    },
    'body': {
      'heading': {
        'top': getValue_(values, 10),
        'text': getValue_(values, 11),
        'bottom': getValue_(values, 12)
      },
      'img': {
        'top': getValue_(values, 13),
        'title': getValue_(values, 14),
        'width': getValue_(values, 15),
        'src': getValue_(values, 16),
        'link': getValue_(values, 17),
        'bottom': getValue_(values, 18)
      },
      'subheading': {
        'top': getValue_(values, 19),
        'text': getValue_(values, 20),
        'bottom': getValue_(values, 21)
      },
      'box': {
        'top': getValue_(values, 22),
        'text': getValue_(values, 23),
        'bottom': getValue_(values, 24)
      }
    },
    'footer': {
      'staff': {
        'top': getValue_(values, 26),
        'workers': [],
        'bottom': getValue_(values, 30)
      },
      'unsubscribe': getValue_(values, 31)
    }
  };
}

function showMailPopup() {
  var ui = SpreadsheetApp.getUi(),
      ss = SpreadsheetApp.getActive(),
      sheet = ss.getSheetByName('Responses'),
      values = sheet.getRange('D3:D34').getValues(),
      mail = HtmlService.createTemplateFromFile('Mail.html'),
      content = getContentObject_(values),
      names = [
        getValue_(values, 27),
        getValue_(values, 28),
        getValue_(values, 29)
      ];
  
  for (var i = 0; i < names.length; i++) {
    var name = names[i],
        nameParts = name.split(' ');
    
    if (nameParts.length == 2) {
      var person = getStaffObject_(nameParts[0], nameParts[1]);
      
      if (content.footer.unsubscribe.toUpperCase() != person.team.toUpperCase()) {
        var resp = ui.alert('Warning', (person.name + ' is not in ' + content.footer.unsubscribe + '. Do you wish to continue?'), ui.ButtonSet.YES_NO);
        
        if (resp == ui.Button.NO) {
          return;
        }
      }
      
      content.footer.staff.workers.push(person);
    }
  }
  
  content = mergeObjects_(content, getDefaultValues_());
  
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
  
  form.content = getDefaultValues_();
  
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

function getDefaultValues_() {
  var ss = SpreadsheetApp.getActive(),
      sheet = ss.getSheetByName('Defaults'),
      values = sheet.getRange('D3:D34').getValues(),
      content = getContentObject_(values);
  
  return content;
}

function getStaffObject_(firstname, lastname) {
  var staffSheetId = Config.get('STAFF_DATA_GSHEET_ID'),
      ss = SpreadsheetApp.openById(staffSheetId),
      sheet = ss.getSheetByName('Staff'),
      values = sheet.getDataRange().getValues(),
      person = {
        'name': firstname + ' ' + lastname,
        'title': '',
        'team': '',
        'photo': getStaffImage_(firstname, lastname)
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

function getStaffImage_(firstname, lastname) {
  var folders = DriveApp.getFoldersByName(lastname + ', ' + firstname),
      imgFile = searchFileInFolder_(folders, 'BubbleHead');
  
  if (imgFile != null) {
    var fileId = imgFile.getId();
    
    return ('https://drive.google.com/uc?export=view&id=' + fileId);
  }
  
  return 'https://vignette.wikia.nocookie.net/citrus/images/6/60/No_Image_Available.png';
}

function searchFileInFolder_(folders, filename) {
  while (folders.hasNext()) {
    var folder = folders.next(),
        files = folder.searchFiles('title contains "' + filename + '"');
    
    if (files.hasNext()) {
      return files.next();
    }
    
    var subfolders = folder.getFolders();
    
    if (subfolders.hasNext()) {
      var file = searchFileInFolder_(subfolders, filename);
      
      if (file != null) {
        return file;
      }
    }
  }
  
  return;
}

function mergeObjects_(obj, src) {
  for (var key in obj) {
    if (obj[key].constructor == Object) {
      obj[key] = mergeObjects_(obj[key], src[key]);
    } else if (String(obj[key]) == '') {
      obj[key] = src[key];
    }
  }
  
  return obj;
}

// redevelopment note: only hide the column if the date in the first Row has lapsed

// move column D to the end of the sheet and hide onOpen
function moveColumn_(iniCol, finCol) {  
  
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