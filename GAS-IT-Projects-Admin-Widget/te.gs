function all(){
  var re = /id=(.*)", "/;
  var plan = SpreadsheetApp.openById(DB).getSheetByName('Requests');
  var allData = plan.getRange('A6:B').getDisplayValues();
  var rawData = plan.getRange('A6:A').getFormulas();
  for(var i in allData){
    var row = allData[i];
    if (row[0] == '') continue;
    var status = row[1];
    if(status == 'IT MANAGER APPROVAL'){
      var col1 = rawData[i][0];
      var id = re.exec(col1)[1];
      //Logger.log('id: %s', id);
      Logger.log('id: %s, val: %s ',id, SpreadsheetApp.openById(id).getSheets()[0].getRange('B11').getValue());
//      var cell = SpreadsheetApp.openById(id).getSheets()[0].getRange('B11');
      //if(cell.getValue() == 4) cell.setValue(10);
    }
  }
}
