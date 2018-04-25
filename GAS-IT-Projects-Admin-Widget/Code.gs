//var DB = '1Oy8ff12OsA4pu45c4mOfmNIcYtRmYt1yK7qHdPlHnk8'; //DEV DB
var DB = '17JKCj7gFAC4veKp7glbO3jtLwLCNbql7URzTAo8z1l4'; // PROD DB
var ARCHIVED_DB = '1e1gsjyPTt-lRIIex-rb6VGeIrM3JAE6I_0yH169Dln8';
//var APP_URL = 'https://script.google.com/a/embraco.com/macros/s/AKfycbySJLQJcTXxfG_93Y31wDXzG8L2-MH2SjRSdQAD5jTNQZ1aYZw/exec'; //DEV APP
var APP_URL = 'https://script.google.com/a/macros/embraco.com/s/AKfycbySNx1efeFJhlyPFzWDW5GPf3KFxLNket8U2KXDZxjH0WWA_c8/exec';  //PROD APP
var EXPORT_FOLDER_ID = '0B973TWs84ZCEcllkSGJ2OTg0aU0';
var PER_PAGE    = 25;
var DATA_RANGE  = 'A6:V';
var USER_DBCACHE = '1q81l8lHUU-9UqpWK8MjKXSyZgh7N7rWpiExAysAYulw';

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

function getTotalLength(){
  var plan = SpreadsheetApp.openById(DB).getSheetByName('Requests');
  var last = plan.getLastRow();
  if(last > 5){
    var total = Number(plan.getRange(6, 1, (plan.getLastRow()-Number(5))).getValues().length);
    var last = Math.ceil(Number(total/PER_PAGE));
    return {total:total, last:last};
  } else {
    return {total:0, last:1};
  }
}

function allDataOnGoingDB(){
  var re = /id=(.*)", "/;
  var user = Session.getActiveUser().getEmail();
  var plan = SpreadsheetApp.openById(DB).getSheetByName('Requests');
  var allData = plan.getRange(DATA_RANGE).getDisplayValues();
  var rawData = plan.getRange('A6:A').getFormulas();
  var data = [];
  for(var i in allData){
    var row = allData[i];
    if (row[0] == '') continue;
    var status = row[1];
    var requester = row[4];
    var manager = row[5];
    var col1 = rawData[i][0];
    var id = re.exec(col1)[1];
    data.push(
      {id:row[0],
       status:status,
       currentResponsible:row[2],
       startDate:row[3],
       requester:requester,
       manager:manager,
       title:row[6],
       department:row[7],
       plant:row[8],
       url:APP_URL+'?id='+id,
       lastUpdate:row[9],
       managerApproval:row[10],
       localITManager:row[11],
       localITManagerApproval:row[12],
       brmManager:row[13],
       brmManagerSend:row[14],
       brm:row[15],
       brmSend:row[16],
       itPortfolioManager:row[17],
       itPortfolioManagerApproval:row[18],
       projectManager:row[19],
       sponsor:row[20],
       evaluation:row[21]!=''?JSON.parse(row[21]):''
      });
  }
  return data;
}

function dataOnGoingDB(page){
  Logger.log(page);
  var re = /id=(.*)"/;
  var user = Session.getActiveUser().getEmail();
  var plan = SpreadsheetApp.openById(DB).getSheetByName('Requests');
  var startLine = Number(page * PER_PAGE)+Number(6)-Number(PER_PAGE);
  var endLine = Number(page * PER_PAGE)+Number(5);
  var allData = plan.getRange('A'+startLine+':J'+endLine).getDisplayValues();
  var rawData = plan.getRange('A'+startLine+':A').getFormulas();
  var data = [];
  for(var i in allData){
    var row = allData[i];
    if (row[0] == '') continue;
    var status = row[1];
    var requester = row[4];
    var manager = row[5];
    var col1 = rawData[i][0];
    var id = re.exec(col1)[1];
    data.push({id:row[0], status:status, currentResponsible:row[2], startDate:row[3], requester:requester, manager:manager, title:row[6],  department:row[7], plant:row[8], url:APP_URL+'?id='+id, lastUpdate:row[9]});
  }
  var totalLast = getTotalLength();
  var data = {data:data, total:totalLast.total, last:totalLast.last }
  return data;
}

function allDataFromArchive(){
  var re = /id=(.*)"/;
  var user = Session.getActiveUser().getEmail();
  var allData = [];
  var rawData = [];

  var plans = SpreadsheetApp.openById(ARCHIVED_DB).getSheets();
  plans.forEach(function (current){
    var newData = current.getRange(DATA_RANGE).getDisplayValues().filter(function (current){
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
    var col1 = rawData[i][0];
    var id = re.exec(col1)[1];
    data.push({id:row[0], status:status, currentResponsible:row[2], startDate:row[3], requester:requester, manager:manager, title:row[6],  department:row[7], plant:row[8], url:APP_URL+'?id='+id, lastUpdate:row[9]});
  }
  return data;
}

function searchRequest(filter, page){
//  filter = {requester:'', manager:'', portfolio:'', endDate:'', plant:'', id:'', title:'', brm:'', localIT:'', pm:'', startDate:'', status:''};
  //page = 1;
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
  if(filter.plant != ''){
    data = data.filter(function(current){
      return current.plant == filter.plant
    });
  }

  if(filter.requester != ''){
    data = data.filter(function(current){
      return current.requester == filter.requester
    });
  }
  if(filter.manager != ''){
    data = data.filter(function(current){
      return current.manager == filter.manager
    });
  }
  if(filter.localIT != ''){
    data = data.filter(function(current){
      return current.localITManager == filter.localIT
    });
  }
  if(filter.brm != ''){
    data = data.filter(function(current){
      return current.brm == filter.brm
    });
  }
  if(filter.portfolio != ''){
    data = data.filter(function(current){
      return current.itPortfolioManager == filter.portfolio
    });
  }
  if(filter.pm != ''){
    data = data.filter(function(current){
      return current.projectManager == filter.pm
    });
  }

  if(filter.startDate != '' && filter.endDate != ''){
    data = data.filter(function(current){
      Logger.log(current.startDate);
      Logger.log(filter.startDate);
      Logger.log(filter.endDate);
       return moment(current.startDate).isBetween(filter.startDate, filter.endDate, 'day', '[]');
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
  var titles = ['REQUEST ID', 'STATUS', 'CURRENTLY WITH', 'START DATE', 'REQUESTER', 'LOCAL AREA MANAGER', 'TITLE', 'DEPARTMENT', 'PLANT', 'LAST UPDATE', 'LOCAL AREA MANAGER APPROVAL', 'IT PROJECT PRIORITIZATION FORUM', 'IT PROJECT PRIORITIZATION FORUM APPROVAL', 'IT MANAGER', 'IT MANAGER SEND', 'IT PROJECT LEADER ANALYSIS', 'IT PROJECT LEADER ANALYSIS SEND', 'IT PORTFOLIO MANAGER', 'IT PORTFOLIO MANAGER APPROVAL', 'IT PROJECT EXECUTION', 'SPONSOR', 'EVALUATION'];
  var export = SpreadsheetApp.create('[IT Projects] Exported Requests - '+Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd HH:mm"), data.length, titles.length);
  var newSheet = DriveApp.getFileById(export.getId());
  DriveApp.getFolderById(EXPORT_FOLDER_ID).addFile(newSheet);
  DriveApp.getRootFolder().removeFile(newSheet);
  var newdata = [];
  newdata.push(titles);
  for(var i in data){
    var row = data[i];
    newdata.push([row.id, row.status, row.currentResponsible, row.startDate, row.requester, row.manager, row.title, row.department, row.plant, row.lastUpdate, row.managerApproval, row.localITManager, row.localITManagerApproval, row.brmManager, row.brmManagerSend, row.brm, row.brmSend, row.itPortfolioManager, row.itPortfolioManagerApproval, row.projectManager, row.sponsor, row.evaluation]);
  }
  export.getActiveSheet().setName('Requests');
  export.getActiveSheet().getRange(1, 1, newdata.length, titles.length).setValues(newdata);
  export.addEditor(Session.getActiveUser().getEmail());
  return export.getUrl();
}
