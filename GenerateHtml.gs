function generateHtmlEmail_() {
  var ss = SpreadsheetApp.getActive();
  var activeRange = SpreadsheetApp.getActiveRange();
  var activeSheet = SpreadsheetApp.getActiveSheet();
  
  if (activeSheet.getName() !== 'Input') {
    throw new Error('Select a cell/range in the "Input" tab.')
  }
  
  var activeColumnNumber = activeRange.getColumn();
  
  if (activeColumnNumber < 6) {
    throw new Error('Select a cell/range in column F or higher.');
  }
  
  var numberOfRows = activeSheet.getLastRow() - 3 - 4 - 2 // - headers - old opt-out - 2 unused 
  
  if (numberOfRows !== 64) {
    throw new Error('There should be 64 rows for style config, found: ' + numberOfRows);
  }
  
  var values = activeSheet.getRange(4, activeColumnNumber, numberOfRows, 1).getValues();  
  var content = getContentObject(values)
  
  var names = [
    getValue(values, 60),
    getValue(values, 61),
    getValue(values, 62)
  ];
      
  for (var i = 0; i < names.length; i++) {
    var name = names[i];
    
    if (name === '') {
      continue;
    }
    
    var nameParts = name.split(' ');

    if (nameParts.length === 2) {
      var person = getStaffObject(nameParts[0], nameParts[1]);      
      content.footer.staff.workers.push(person);     
    } else {
      throw new Error('Too many/few parts in the name: "' + name + '"');
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

  var mail = HtmlService.createTemplateFromFile('Mail.html');
  mail.content = content;
  var html = mail.evaluate().setWidth(800).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Generated mail');
  return;
  
  // Private Functions
  // -----------------
  
  function getDefaultValues() {
    var sheet = ss.getSheetByName('Defaults');
    var values = sheet.getRange('F4:F68').getValues();
    var content = getContentObject(values);
    return content;
  }
  
  function getStaffObject(firstname, lastname) {
    var staffSheetId = Config.get('STAFF_DATA_GSHEET_ID');
    var ss = SpreadsheetApp.openById(staffSheetId);
    var sheet = ss.getSheetByName('Staff Directory');
    var values = sheet.getDataRange().getValues();
    var person = {
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
    var folders = DriveApp.getFoldersByName(lastname + ', ' + firstname);
    var imgFile = searchFileInFolder(folders, 'BubbleHead');
    
    if (imgFile != null) {
      var fileId = imgFile.getId();      
      return ('https://drive.google.com/uc?export=view&id=' + fileId);
    }
    
    return 'https://vignette.wikia.nocookie.net/citrus/images/6/60/No_Image_Available.png';
    
    // Private Functions
    // -----------------
    
    function searchFileInFolder(folders, filename) {
      while (folders.hasNext()) {
        var folder = folders.next();
        var files = folder.searchFiles('title contains "' + filename + '"');
        
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
    
  } // generateHtmlEmail_.searchFileInFolder()
  
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
        }
      }
    };
  }
  
  function getValue(values, index) {
    return String(values[index][0]).trim();
  }  

} // generateHtmlEmail_()

