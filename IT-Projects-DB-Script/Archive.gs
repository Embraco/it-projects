function archiveOldRequests(){
  var currentYear = new Date().getFullYear();
  var archiveDB = SpreadsheetApp.openById(ARCHIVE_DB);
  var archivePlan = archiveDB.getSheetByName(currentYear);
  if(archivePlan == null){
    archivePlan = archiveDB.getSheetByName('TEMPLATE').copyTo(archiveDB);
    archivePlan.setName(currentYear);
    archivePlan.showSheet();
  }
  var requestsPlan = SpreadsheetApp.getActive().getSheetByName('Requests');
  var rows = requestsPlan.getRange('B6:J').getValues();
  var formulas = requestsPlan.getRange('A6:A').getFormulas();
  var finishedProjects = [];
  for(var index in rows){
    var currentValue = rows[index];
    if(currentValue[0] == 'FINISHED'){
      var n = formulas[index].concat(currentValue);
      finishedProjects.push(n);
    }
  };
  var continuee = false;
  try{
    archivePlan.getRange(archivePlan.getLastRow()+1, 1, finishedProjects.length, finishedProjects[0].length).setValues(finishedProjects);
    continuee = true;
  } catch(e){
    SpreadsheetApp.getActive().toast('Something goes wrong on Archive process. Please contact the administrator.');
  }
  if(continuee==true){
    for(var index = Number(rows.length-1);index >= 0; index--){
      var currentValue = rows[index];
      if(currentValue[0] == 'FINISHED'){
        requestsPlan.deleteRow(Number(index)+Number(6));
      }
    };
  }
  SpreadsheetApp.getActive().toast('Archive process successfully finished.');
}
