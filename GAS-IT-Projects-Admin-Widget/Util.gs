function usersCache(){
  var plan = SpreadsheetApp.openById(USER_DBCACHE);
  var sheet = plan.getSheetByName('Users');
  var users = [];
  var values = sheet.getRange('A2:B').getValues();
  for(var i in values){
    if (values[i][0] == '') continue;
    users.push({text:values[i][0], value:values[i][1]});
  }
  return users;
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
