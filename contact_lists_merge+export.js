/**
 * @OnlyCurrentDoc  Limits the script to only accessing the current spreadsheet.
 */

/**
 * Adds a custom menu 
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */

String.prototype.capitalize = function(lower) {
    return (lower ? this.toLowerCase() : this).replace(/(?:^|\s)\S/g, function(a) { return a.toUpperCase(); });
};

function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Update')
      .addItem('Update Resp.ALL + csv', 'updateTab')
      .addToUi();
}

/**
 * Runs when the add-on is installed; calls onOpen() to ensure menu creation and
 * any other initializion work is done immediately.
 *
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
  onOpen(e);
}

function updateTab() {
  var sheet_update = SpreadsheetApp.getActiveSpreadsheet();
  sheet_update.insertSheet('All as ' + Date(), 14);
  var sheet_update = sheet_update.getSheets()[14];
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet.setActiveSheet(sheet.getSheets()[1]);
  var data = sheet.getDataRange().getValues();
  sheet_update.appendRow([data[1][0],data[1][1],data[1][2],data[1][3],data[1][4],data[1][5],data[1][6],"'" + data[1][7],"'" + data[1][8],data[1][9]]);
  sheet_update.setFrozenRows(1);
  
  var cell = sheet_update.getRange("A1:J1");               
   cell.setFontColor('white');                     
   cell.setBackground('red');    
  
  var sheet_csv = SpreadsheetApp.getActiveSpreadsheet();
  sheet_csv.insertSheet('csv as ' + Date(), 15);
  var sheet_csv = sheet_csv.getSheets()[15];

  sheet_csv.appendRow(['Name','Occupation','Group Membership','E-mail 1 - Type','E-mail 1 - Value','Phone 1 - Type','Phone 1 - Value','Notes']);
  sheet_csv.setFrozenRows(1);
  
  var cell = sheet_update.getRange("A1:J1");               
   cell.setFontColor('white');                     
   cell.setBackground('red');
  
  for (var i = 1; i < 11; i++) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    sheet.setActiveSheet(sheet.getSheets()[i]);
    var data = sheet.getDataRange().getValues();
      for (var j = 2; j < data.length; j++) {
        sheet_update.appendRow([data[j][0],data[j][1],data[j][2],data[j][3],data[j][4],data[j][5],data[j][6],"'" + data[j][7],"'" + data[j][8],data[j][9]]);
        
        phone1 = data[j][7];
        if (typeof phone1 == "number"){
          phone1 = String(phone1);
        }
        phone1 = phone1.replace(/[^0-9]/gi,"");
        if (phone1.substr(0,3) == "977"){
          phone1 = phone1.substr(3,15);
        }
        //phone1 = String(parseInt(phone1));
        phone1 = phone1.substr(0,3) + " " + phone1.substr(3,3) + " " + phone1.substr(6,4) + " " + phone1.substr(10,5);

        phone2 = String(data[j][8]);
        if (typeof phone2 == "number"){
          phone2 = String(phone2);
        }
        phone2 = phone2.replace(/[^0-9]/gi,"");
        if (phone2.substr(0,3) == "977"){
          phone2 = phone2.substr(3,15);
        }
        //phone2 = String(parseInt(phone2));
        phone2 = phone2.substr(0,3) + " " + phone2.substr(3,3) + " " + phone2.substr(6,4) + " " + phone2.substr(10,5);
        
        if (data[j][8]){
           sheet_csv.appendRow([data[j][5].capitalize(true),data[j][9],sheet.getSheetName(),'Work',data[j][6],'Local',phone1,"(2) local phone numbers"]);
           sheet_csv.appendRow([data[j][5].capitalize(true),data[j][9],sheet.getSheetName(),'Work',data[j][6],'Local',phone2,"(2) local phone numbers"]);
        }else{
           if (data[j][7]){
             sheet_csv.appendRow([data[j][5].capitalize(true),data[j][9],sheet.getSheetName(),'Work',data[j][6],'Local',phone1,"(1) local phone number"]);
           }else{
             sheet_csv.appendRow([data[j][5].capitalize(true),data[j][9],sheet.getSheetName(),'Work',data[j][6],'Local',"'" + data[j][7],"(0) local phone number"]);
           }
        }
      }
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet.setActiveSheet(sheet.getSheets()[14]);
}
