//var DB = '1Oy8ff12OsA4pu45c4mOfmNIcYtRmYt1yK7qHdPlHnk8'; //DEV DB
var DB = '17JKCj7gFAC4veKp7glbO3jtLwLCNbql7URzTAo8z1l4'; // PROD DB
//var APP_URL = 'https://script.google.com/a/embraco.com/macros/s/AKfycbySJLQJcTXxfG_93Y31wDXzG8L2-MH2SjRSdQAD5jTNQZ1aYZw/exec'; //DEV APP
var APP_URL = 'https://script.google.com/a/macros/embraco.com/s/AKfycbySNx1efeFJhlyPFzWDW5GPf3KFxLNket8U2KXDZxjH0WWA_c8/exec';  //PROD APP;
var ARCHIVED_DB = '1e1gsjyPTt-lRIIex-rb6VGeIrM3JAE6I_0yH169Dln8';
var EXPORT_FOLDER_ID = '0B973TWs84ZCEcllkSGJ2OTg0aU0';
var PER_PAGE    = 25;

function doGet(e) {
  var html = HtmlService.createTemplateFromFile('index');

  return html.evaluate()
  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setTitle("Embraco - IT Project Requests")
  .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, minimum-scale=1.0');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function allDataFromArchive(){
  var re = /id=(.*)", "/;
  var user = Session.getActiveUser().getEmail();
  var allData = [];
  var rawData = [];

  var plans = SpreadsheetApp.openById(ARCHIVED_DB).getSheets();
  plans.forEach(function (current){
    var newData = current.getRange('A6:J').getDisplayValues().filter(function (current){
       return current[0] != ''
    });
    var newRawData = current.getRange('A6:A').getFormulas().filter(function (current){
       return current[0] != ''
    });
    rawData = rawData.concat(newRawData);
    allData = allData.concat(newData);
  });

  var data = [];
  for(var i in allData){
    var row = allData[i];
    if(row[0] == '') continue;
    var status = row[1];
    var requester = row[4];
    var manager = row[5];
    var currentResponsible = row[2];
    if (user == requester || user == manager || user == currentResponsible){
      var col1 = rawData[i][0];
      var id = re.exec(col1)[1];
      data.push({id:row[0], status:status, currentResponsible:row[2], startDate:row[3], requester:requester, manager:manager, title:row[6],  department:row[7], plant:row[8], url:APP_URL+'?id='+id, lastUpdate:row[9]});
    }
  }
  return data;
}

function allDataOnGoingDB(){
  var re = /id=(.*)"/;
  var user = Session.getActiveUser().getEmail();
  var plan = SpreadsheetApp.openById(DB).getSheetByName('Requests');
  var allData = plan.getRange('A6:J').getDisplayValues();
  var rawData = plan.getRange('A6:A').getFormulas();
  var data = [];
  for(var i in allData){
    var row = allData[i];
    var status = row[1];
    var requester = row[4];
    var manager = row[5];
    var currentResponsible = row[2];
    if (user == requester || user == manager || user == currentResponsible){
      var col1 = rawData[i][0];
      var id = re.exec(col1)[1];
      data.push({id:row[0], status:status, currentResponsible:row[2], startDate:row[3], requester:requester, manager:manager, title:row[6],  department:row[7], plant:row[8], url:APP_URL+'?id='+id, lastUpdate:row[9]});
    }
  }
  return data;
}

function dataOnGoingDB(page){
  Logger.log(page);
  var re = /id=(.*)"/;
  var user = Session.getActiveUser().getEmail();
  var plan = SpreadsheetApp.openById(DB).getSheetByName('Requests');
  var allData = plan.getRange('A6:J').getDisplayValues();
  var rawData = plan.getRange('A6:A').getFormulas();
  var data = [];
  for(var i in allData){
    var row = allData[i];
    if (row[0] == '') continue;
    var status = row[1];
    var requester = row[4];
    var manager = row[5];
    var currentResponsible = row[2];
    if (user == requester || user == manager || user == currentResponsible){
      var col1 = rawData[i][0];
      var id = re.exec(col1)[1];
      data.push({id:row[0], status:status, currentResponsible:row[2], startDate:row[3], requester:requester, manager:manager, title:row[6],  department:row[7], plant:row[8], url:APP_URL+'?id='+id, lastUpdate:row[9]});
    }
  }
  var total = data.length;
  var startLine = Number(page * PER_PAGE)-Number(PER_PAGE);
  var endLine = Number(page * PER_PAGE);
  data = data.slice(startLine, endLine);

  var last = Math.ceil(Number(total/PER_PAGE));
  var data = {data:data, total:total, last:last }
  return data;
}

function searchRequest(filter, page){
  Logger.log(filter);
  var data = allDataOnGoingDB();

  if (filter.id != ''){
    for (var i in data){
      var row = data[i];
      if(row.id == filter.id) {
         return {data:[row], total:1, last:1};
      }
    }
    // if ID is found in current database, retrieve the archive database to search.
    data = data.concat(allDataFromArchive());
    for (var i in data){
      var row = data[i];
      if(row.id == filter.id) {
         return {data:[row], total:1, last:1};
      }
    }
    return {data:0, total:0, last:0};
  }
  if(filter.status == 'FINISHED'){
    data = data.concat(allDataFromArchive());
  }
  // filter by ID stop the flow and return only one
  if (filter.title != ''){
    data = data.filter(function(current){
      return current.title.indexOf(filter.title) > -1
    });
  }
  if(filter.status != ''){
    data = data.filter(function(current){
      return current.status == filter.status
    });
  }
  var total = data.length;
  if(page != 0){
    var last = Math.ceil(Number(total/PER_PAGE));
    var startLine = Number(page * PER_PAGE)-Number(PER_PAGE);
    var endLine = Number(page * PER_PAGE);
    data = data.slice(startLine, endLine);
    Logger.log('Start: %s - End: %s', startLine, endLine);
  }
  return {data:data, total:total, last:last};
}

function exportToSheets(filter){
  var data = searchRequest(filter, 0).data;
  var export = SpreadsheetApp.create('[IT Projects] Exported Requests - '+Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd HH:mm"), data.length, 10);
  var newSheet = DriveApp.getFileById(export.getId());
  DriveApp.getFolderById(EXPORT_FOLDER_ID).addFile(newSheet);
  DriveApp.getRootFolder().removeFile(newSheet);
  var newdata = [];
  newdata.push(['ID', 'STATUS', 'CURRENT WITH', 'START DATE', 'REQUESTER', 'LOCAL AREA MANAGER', 'TITLE', 'DEPARTMENT', 'PLANT','LAST UPDATE'])
  for(var i in data){
    var row = data[i];
    newdata.push([row.id,row.status,row.currentResponsible,row.startDate,row.requester,row.manager,row.title,row.department,row.plant,row.lastUpdate]);
  }
  export.getActiveSheet().getRange(1, 1, newdata.length, 10).setValues(newdata);
  export.addEditor(Session.getActiveUser().getEmail());
  return export.getUrl();
}
