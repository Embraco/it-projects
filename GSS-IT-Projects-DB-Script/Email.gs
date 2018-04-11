var STATUS_CELL = 'B11';
var ARCHIVE_DB = '1e1gsjyPTt-lRIIex-rb6VGeIrM3JAE6I_0yH169Dln8';
var USER_DBCACHE = '1q81l8lHUU-9UqpWK8MjKXSyZgh7N7rWpiExAysAYulw';
var APP_URL = 'https://script.google.com/a/macros/embraco.com/s/AKfycbySNx1efeFJhlyPFzWDW5GPf3KFxLNket8U2KXDZxjH0WWA_c8/exec';
var projectStatus = {
  DRAFT: 1,
  BUSINESS_APPROVAL: 2,
  IT_APPROVAL: 3,
  BRM_MANAGERS: 10,
  PROJECT_ANALYSIS: 4,
  PROJECT_PRIORITIZATION_FOR_EXECUTION: 5,
  PROJECT_EXECUTION: 6,
  PROJECT_IN_EVALUATION:7,
  FINISHED: 8,
  CANCELED: 9
};

function onOpen(){
  var ui = SpreadsheetApp.getUi();
  var plan_name = SpreadsheetApp.getActiveSpreadsheet().getName();
  var menu = ui.createMenu('Embraco');
  menu.addItem('Archive old Requests', 'archiveOldRequests');
  menu.addSeparator()
  menu.addItem('Change Local Area Manager', 'changeManager');
  menu.addItem('Change Local IT Manager', 'changeITLocalManager');
  menu.addItem('Change IT Manager Approval', 'changeBRM');
  menu.addItem('Change IT Project Leader Analysis', 'changeBRMAnalyst');
  menu.addItem('Change IT Portfolio Manager', 'changePortfolio');
  menu.addItem('Change Project Manager', 'changeProjectManager');
  menu.addToUi();
}

//===========================================================================================================================
//===========================================================================================================================
function usersCache(){
  var plan = SpreadsheetApp.openById(USER_DBCACHE);
  var sheet = plan.getSheetByName('Users');
  var users = [];
  var values = sheet.getRange('A2:B').getValues();
  for(var i in values){
    users.push({text:values[i][0], value:values[i][1]});
  }
  return users;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
