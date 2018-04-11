function changeManager() {
  var active = SpreadsheetApp.getActiveSheet();
  var row = active.getActiveCell().getRow();
  if(active.getRange(row, 2).getValue() == 'LOCAL AREA MANAGER APPROVAL')
    changeCore('requesterManager', 'B4');
  else
    SpreadsheetApp.getUi().alert('Alert', 'Unable to change a step already approved!', SpreadsheetApp.getUi().ButtonSet.OK);
}
function changeITLocalManager() {
  var active = SpreadsheetApp.getActiveSheet();
  var row = active.getActiveCell().getRow();
  if(active.getRange(row, 2).getValue() == 'IT PROJECT PRIORITIZATION FORUM')
  changeCore('projectLocalITManager', 'B16');
  else
    SpreadsheetApp.getUi().alert('Alert', 'Unable to change a step already approved!', SpreadsheetApp.getUi().ButtonSet.OK);
}
function changeBRM() {
  var active = SpreadsheetApp.getActiveSheet();
  var row = active.getActiveCell().getRow();
  if(active.getRange(row, 2).getValue() == 'IT MANAGER APPROVAL')
  changeCore('projectBrm', 'B21');
  else
    SpreadsheetApp.getUi().alert('Alert', 'Unable to change a step already approved!', SpreadsheetApp.getUi().ButtonSet.OK);
}
function changeBRMAnalyst() {
  var active = SpreadsheetApp.getActiveSheet();
  var row = active.getActiveCell().getRow();
  if(active.getRange(row, 2).getValue() == 'BRM ANALYST')
  changeCore('projectBrmManagerBrmAnalyst', 'D17');
  else
    SpreadsheetApp.getUi().alert('Alert', 'Unable to change a step already approved!', SpreadsheetApp.getUi().ButtonSet.OK);
}
function changePortfolio() {
  var active = SpreadsheetApp.getActiveSheet();
  var row = active.getActiveCell().getRow();
  if(active.getRange(row, 2).getValue() == 'IT PORTFOLIO MANAGER APPROVAL')
  changeCore('projectBrmPortfolioManager', 'B26');
  else
    SpreadsheetApp.getUi().alert('Alert', 'Unable to change a step already approved!', SpreadsheetApp.getUi().ButtonSet.OK);
}
function changeProjectManager() {
  var active = SpreadsheetApp.getActiveSheet();
  var row = active.getActiveCell().getRow();
  if(active.getRange(row, 2).getValue() == 'PROJECT MANAGER')
  changeCore('projectPm', 'B31');
  else
    SpreadsheetApp.getUi().alert('Alert', 'Unable to change a step already approved!', SpreadsheetApp.getUi().ButtonSet.OK);
}

function changeCore(manager, location){
  if(SpreadsheetApp.getActiveSheet().getName() == 'Requests'){
    var activeRowNumber = SpreadsheetApp.getActiveSheet().getActiveCell().getRow();
    var id = SpreadsheetApp.getActiveSheet().getRange(activeRowNumber, 1).getFormula();
    var idRaw = SpreadsheetApp.getActiveSheet().getRange(activeRowNumber, 1).getDisplayValue();
    var regexp = /id=(.*)\"\,/
    var result = regexp.exec(id);
    var requestPlanId = result[1];
    var html = HtmlService.createTemplateFromFile('popup');
    html.requestPlanId = requestPlanId;
    html.fieldToChange = manager;
    html.requestLocation = location;
    var eva = html.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setHeight(350).setWidth(600);
    SpreadsheetApp.getUi().showModalDialog(eva, 'Change the Request #'+idRaw);
  } else{
    SpreadsheetApp.getUi().alert('You should be in Requests sheet to be able to execute this action.');
  }
}

function saveChange(requestPlanId, newValue, location, type){
  try{
    var active = SpreadsheetApp.getActiveSheet();
    var row = active.getActiveCell().getRow();
    var sheet = SpreadsheetApp.openById(requestPlanId).getSheets()[0];

    sheet.getRange(location).setValue(newValue);

    if(type == 'requesterManager'){
      active.getRange(row, 6).setValue(newValue);
    }
    active.getRange(row, 3).setValue(newValue);

    var id = sheet.getRange('B9').getValue();
    var projectName = sheet.getRange('B10').getValue();
    var requesterEmail = sheet.getRange('B3').getValue();
    sendEmail('ChangeOwnerEmail', newValue, 'Request '+id+' needs your attention', requestPlanId, projectName, requesterEmail, '');
    return {answer:'OK', error:''};
  } catch(e){
    return {answer:'OK', error:e.message + ' ('+e.lineNumber+')'};
  }
}

function getStatusString(status){
  switch (status) {
        case projectStatus.DRAFT:
            return 'DRAFT';
        case projectStatus.BUSINESS_APPROVAL:
            return 'LOCAL AREA MANAGER APPROVAL';
        case projectStatus.IT_APPROVAL:
            return 'IT PROJECT PRIORITIZATION FORUM';
        case projectStatus.BRM_MANAGERS:
            return 'IT MANAGER APPROVAL';
        case projectStatus.PROJECT_ANALYSIS:
            return 'BRM ANALYST';
        case projectStatus.PROJECT_PRIORITIZATION_FOR_EXECUTION:
            return 'IT PORTFOLIO MANAGER APPROVAL';
        case projectStatus.PROJECT_EXECUTION:
            return 'PROJECT MANAGER';
        case projectStatus.PROJECT_IN_EVALUATION:
            return 'PROJECT IN EVALUATION';
        case projectStatus.FINISHED:
            return 'FINISHED';
        case projectStatus.CANCELED:
            return 'CANCELED';
    }
}
