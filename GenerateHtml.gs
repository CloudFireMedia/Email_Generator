function generateHtmlEmail_() {
  var ui = SpreadsheetApp.getUi(),
      ss = SpreadsheetApp.getActive(),
      sheet = ss.getSheetByName('Input'),
      values = sheet.getRange('F4:F68').getValues(),
      mail = HtmlService.createTemplateFromFile('Mail.html'),
      content = getContentObject_(values),
      names = [
        getValue_(values, 60),
        getValue_(values, 61),
        getValue_(values, 62)
      ];
  
  for (var i = 0; i < names.length; i++) {
    var name = names[i],
        nameParts = name.split(' ');
    
    if (nameParts.length == 2) {
      var person = getStaffObject(nameParts[0], nameParts[1]);      
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
  return;
  
  // Private Functions
  // -----------------
  
  function getDefaultValues() {
    var ss = SpreadsheetApp.getActive(),
        sheet = ss.getSheetByName('Defaults'),
        values = sheet.getRange('F4:F68').getValues(),
        content = getContentObject_(values);
    
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
    
    // Private Functions
    // -----------------
    
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

} // generateHtmlEmail_()

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
      'section_1_text': {
        'top': getValue_(values, 10),
        'text': getValue_(values, 11),
        'bottom': getValue_(values, 12)
      },
      'section_1_box': {
        'top': getValue_(values,13),
        'text': getValue_(values, 14),
        'bottom': getValue_(values, 15)
      },
      'section_1_img': {
        'top': getValue_(values, 16),
        'title': getValue_(values, 17),
        'width': getValue_(values, 18),
        'src': getValue_(values, 19),
        'link': getValue_(values, 20),
        'bottom': getValue_(values, 21)
      },
      'section_2_text': {
        'top': getValue_(values, 22),
        'text': getValue_(values, 23),
        'bottom': getValue_(values, 24)
      },
      'section_2_box': {
        'top': getValue_(values, 25),
        'text': getValue_(values, 26),
        'bottom': getValue_(values, 27)
      },
      'section_2_img': {
        'top': getValue_(values, 28),
        'title': getValue_(values, 29),
        'width': getValue_(values, 30),
        'src': getValue_(values, 31),
        'link': getValue_(values, 32),
        'bottom': getValue_(values, 33)
      },
      'section_3_text': {
        'top': getValue_(values, 34),
        'text': getValue_(values, 35),
        'bottom': getValue_(values, 36)
      },
      'section_3_box': {
        'top': getValue_(values, 37),
        'text': getValue_(values, 38),
        'bottom': getValue_(values, 39)
      },
      'section_3_img': {
        'top': getValue_(values, 40),
        'title': getValue_(values, 41),
        'width': getValue_(values, 42),
        'src': getValue_(values, 43),
        'link': getValue_(values, 44),
        'bottom': getValue_(values, 45)
      },
      'section_4_text': {
        'top': getValue_(values, 46),
        'text': getValue_(values, 47),
        'bottom': getValue_(values, 48)
      },
      'section_4_box': {
        'top': getValue_(values, 49),
        'text': getValue_(values, 50),
        'bottom': getValue_(values, 51)
      },
      'section_4_img': {
        'top': getValue_(values, 52),
        'title': getValue_(values, 53),
        'width': getValue_(values, 54),
        'src': getValue_(values, 55),
        'link': getValue_(values, 56),
        'bottom': getValue_(values, 57)
      },
      
    },
    'footer': {
      'staff': {
        'top': getValue_(values, 59),
        'workers': [],
        'bottom': getValue_(values, 63)
      },
      'unsubscribe': getValue_(values, 64)
    }
  };
}

function getValue_(values, index) {
  return String(values[index][0]).trim();
}