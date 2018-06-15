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

 	var range = SpreadsheetApp.getActiveRange();
 	if (range != 'nothing' && range != null) {

 		var sh = range.getSheet();

 		if (sh.getName() == 'Responses') {

 			var row = range.getRow();

 			if (row > 1) {

 				var rowData = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues();

 				if (rowData[0][8] != '') {

 					var e = {
 						values : rowData[0]

 					}

 					onFormSubmit(e);
                    
                    SpreadsheetApp.getActiveSpreadsheet().toast("Successfully processed!","Info", 2);

 				} else {

 					var ui = SpreadsheetApp.getUi();
 					var result = ui.alert('Error', 'Invalid row selected!', ui.ButtonSet.OK);
 					return;

 				}
                
 			} else {

 				var ui = SpreadsheetApp.getUi();
 				var result = ui.alert('Error', 'Invalid selection. You must select a cell within the row that you wish to process.', ui.ButtonSet.OK);
 				return;

 			}
            
 		} else {

 			var ui = SpreadsheetApp.getUi();
 			var result = ui.alert('Error', 'Invalid sheet!', ui.ButtonSet.OK);
 			return;

 		}

 	} else {

 		var ui = SpreadsheetApp.getUi();
 		var result = ui.alert('Error', 'Invalid selection. You must select a cell within the row that you wish to process.', ui.ButtonSet.OK);
 		return;

 	}

 }
 
 function setDefaults(){

var sp = SpreadsheetApp.getActiveSpreadsheet();
var sh=sp.getSheetByName('Default');
var data=sh.getDataRange().getValues();


 var obj={
 "f1":data[1][2],
 "f2":data[2][2],
 "f3":data[3][2],
 "f4":data[4][2],
 "f5":data[5][2],
 "f6":data[6][2],
 "f13":data[8][2],
 "f9":data[10][2],
 "f10":data[11][2],
 "f12":data[13][2],


 }
   var html = HtmlService.createTemplateFromFile('HTML')
   html.data= JSON.stringify(obj)
   var out =  html.evaluate()
           .setWidth(500)
           .setHeight(640);
      
      SpreadsheetApp.getUi() 
      .showModalDialog(out, 'Set Defaults');
 
 }
 
 function writeDef(strObj){
 

 var obj= JSON.parse(strObj)
 
 var sp = SpreadsheetApp.getActiveSpreadsheet();
 var sh=sp.getSheetByName('Default');
        sh.getRange(2, 3).setValue(obj.f1)
        sh.getRange(3, 3).setValue(obj.f2)    
        sh.getRange(4, 3).setValue(obj.f3)
        sh.getRange(5, 3).setValue(obj.f4)  
        sh.getRange(6, 3).setValue(obj.f5)
        sh.getRange(7, 3).setValue(obj.f6)   
        sh.getRange(9, 3).setValue(obj.f13)   
        sh.getRange(11, 3).setValue(obj.f9)
        sh.getRange(12, 3).setValue(obj.f10)            
        sh.getRange(14, 3).setValue(obj.f12)  
         


SpreadsheetApp.flush();

var formID="1BaIsWXctQAPlSVMKSennKRc5wz-j3xETu0sZYGf7aoQ";

    var form= FormApp.openById(formID);
    var items = form.getItems(FormApp.ItemType.TEXT)

    var f2= items[0].asTextItem() 
    f2.setHelpText('Default is "' + obj.f2 +'"')
    
    var f3= items[1].asTextItem() 
    f3.setHelpText('Default is ' + obj.f3 +'') 
    

    var f4= items[2].asTextItem() 
    f4.setHelpText('Default is "' + obj.f4 +'"')
    
    var f5= items[3].asTextItem() 
    f5.setHelpText('Default is "' + obj.f5 +'"') 
    
    var f6= items[4].asTextItem() 
    f6.setHelpText('Default is ' + obj.f6 +'')
    
   var f13= items[5].asTextItem() 
    f13.setHelpText('Default is "' + obj.f13 +'"') 
    
    var f9= items[6].asTextItem() 
    f9.setHelpText('Default is "' + obj.f9 +'"') 
    
    var f10= items[7].asTextItem() 
    f10.setHelpText('Default is ' + obj.f10 +'')
    
    var f12= items[9].asTextItem() 
    f12.setHelpText('Default is "' + obj.f12 +'"')

 }
 
 
 
 
 
 function include(filename) {

	return HtmlService.createHtmlOutputFromFile(filename)
	.getContent();

}