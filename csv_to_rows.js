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
      .addItem('Reformat', 'updateTab')
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
  sheet_update.insertSheet('All as ' + Date(), 1);
  var sheet_update = sheet_update.getSheets()[1];    
  
   var i=0;
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    sheet.setActiveSheet(sheet.getSheets()[i]);
  var col0_cat;
  var col1_district;
  var col3_sector;
  
  var data = sheet.getDataRange().getValues();
      for (var j = 4; j < data.length; j++) {
        if (data[j][0]){
          col0_cat = data[j][0];
        }
        if (data[j][2]){
          col1_district = data[j][2];
        }
          //col2_vdc
        if (data[j][3]){
          col3_sector = data[j][3];
        }
          var col4_pns = data[j][4];
        var vdc_temp = data[j][6];
        
        var vdc_temp = vdc_temp.split(', ');
        
        if (data[j][4]){
          for (var k in vdc_temp){
            
            sheet_update.appendRow([col0_cat,col1_district,vdc_temp[k],col3_sector,col4_pns]);
          }
        }
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet.setActiveSheet(sheet.getSheets()[1]);
}
