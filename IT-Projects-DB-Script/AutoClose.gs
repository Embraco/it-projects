function autoCloseRequestInEvaluation(){
  var requestsToCloseSheet = SpreadsheetApp.getActive().getSheetByName('Requests to close');
  var data = requestsToCloseSheet.getRange('A2:C').getValues();
  var today = new Date();
  for(var i in data){
    var row = data[i];
    if(isSameDay(row[2], today)) {
      var plan;
      try{
        plan = SpreadsheetApp.openById(row[1]);
        var sheet = plan.getSheetByName('Request');
        var projectName = sheet.getRange('B10').getValue();
        var requesterEmail = sheet.getRange('B3').getValue();
        sheet.getRange(STATUS_CELL).setValue(projectStatus.FINISHED);
        //------------------------------------------------------------------
        var comment = {comment:'Request has been FINISHED automaticaly.', date:new Date().toString(), commenter:Session.getActiveUser().getEmail()};
        var comments = plan.getSheetByName('RequestComments');
        comments.appendRow([comment.comment, comment.commenter, comment.date]);
        //--------------------------------------
        // email de notificacao
        sendEmail('AutoCloseEmail', requesterEmail, 'Request '+row[0]+' has been finished automaticaly', row[1], projectName, requesterEmail, '');
        updateProjectOnReport(row[0]);
        requestsToCloseSheet.deleteRow((Number(i)+Number(2)));
      } catch(e){
        Logger.log(e);
        sendEmail('AutoCloseEmail', 'it.apps@embraco.com', '[ERRO] Request has been finished', row[1], projectName, requesterEmail, '');
      }
    }
  }
}
function updateProjectOnReport(id){
  var report = SpreadsheetApp.getActive().getSheetByName('Requests');
  var rows = report.getRange('A6:A').getDisplayValues();
  for(var i in rows) {
    var row = rows[i];
    if (row == '') break;
    var idd = row[0];
    if (idd == id){
      report.getRange((+i+6), 2).setValue('FINISHED');
      break;
    }
  }
}

function isSameDay(date1, date2){
  return moment(date1).isSame(date2, 'day');

}
