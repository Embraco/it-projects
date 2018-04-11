var PROJECT_TEMPLATE_ID = '1tfR61ryhNheoSMEiaR49m_nbXHaftZPHz5ScBpswsgE';
var requestPlansFolder = '0B973TWs84ZCENFRvN1k4bXI4T2c';
var DB = '17JKCj7gFAC4veKp7glbO3jtLwLCNbql7URzTAo8z1l4'; // PROD DB 17JKCj7gFAC4veKp7glbO3jtLwLCNbql7URzTAo8z1l4 -- DEV DB 1Oy8ff12OsA4pu45c4mOfmNIcYtRmYt1yK7qHdPlHnk8
var USER_DBCACHE = '1q81l8lHUU-9UqpWK8MjKXSyZgh7N7rWpiExAysAYulw';
var COMMENT_RANGE = 'A3:C';
var ATTACHMENTS_RANGE = 'A3:F';
var STATUS_CELL = 'B11';
var ID_CELL = 'B9';
var APP_URL = ScriptApp.getService().getUrl();
var LOGS_PLAN = '18i2v2Ili1AgRs2Ei2Pi--NmO8VrvIZc9z5-VLGKTyAI';
var projectStatus = {
    DRAFT: 1,
    BUSINESS_APPROVAL: 2,
    IT_APPROVAL: 3,
    BRM_MANAGERS: 10,
    PROJECT_ANALYSIS: 4,
    PROJECT_PRIORITIZATION_FOR_EXECUTION: 5,
    PROJECT_EXECUTION: 6,
    PROJECT_IN_EVALUATION: 7,
    FINISHED: 8,
    CANCELED: 9
};

function newProject() {
    var np = {
        projectID: 0,
        comments: [],
        projectAttachmets: []
    };
    var plan = SpreadsheetApp.openById(PROJECT_TEMPLATE_ID);
    var values = plan.getRangeByName('ProjectData').getValues();
    for (var i in values) {
        var line = values[i]
        np[line[0]] = line[1];
    }
    var values = plan.getRangeByName('ProjectData2').getValues();
    for (var i in values) {
        var line = values[i]
        np[line[0]] = line[1];
    }
    var values = plan.getSheetByName('Request').getRange('C16:D20').getValues();
    for (var i in values) {
        var line = values[i]
        np[line[0]] = line[1];
    }
    np.projectID = 0;
    var standardITDevelopmentManager = getConfigByKey('StandardITBusinessDevelopmentManager');
    if (standardITDevelopmentManager == undefined) standardITDevelopmentManager = 'luciano.borges@embraco.com';
    np.projectLocalITManager = standardITDevelopmentManager
    var email = Session.getActiveUser().getEmail();
    np = retrieveUserInfo(email, np);
    return np;
}
//-------------------------------------------------------------------------------------------------------
//-------------------------------------------------------------------------------------------------------
function saveProjectNewRequest(np) {
    var plan;
    var comment;
    try {
        var arrayToSheets = [];
        if (np.projectPlanID == '') {
            plan = SpreadsheetApp.openById(PROJECT_TEMPLATE_ID).copy('[DRAFT] New IT Project Request - ' + np.requesterEmail);
            np.projectPlanID = plan.getId();
            var driveFile = DriveApp.getFileById(plan.getId());
            DriveApp.getFolderById(requestPlansFolder).addFile(driveFile);
            DriveApp.getRootFolder().removeFile(driveFile);
        } else {
            plan = SpreadsheetApp.openById(np.projectPlanID);
        }
        arrayToSheets.push([np.requesterName]);
        arrayToSheets.push([np.requesterEmail.toLowerCase()]);
        arrayToSheets.push([np.requesterManager.toLowerCase()]);
        arrayToSheets.push([np.requesterPhone]);
        arrayToSheets.push([np.requesterDepartment]);
        arrayToSheets.push([np.requesterLocation]);
        arrayToSheets.push([np.requestDate]);
        arrayToSheets.push([np.projectID]);
        arrayToSheets.push([np.projectName]);
        arrayToSheets.push([np.projectStatus]);
        arrayToSheets.push([np.projectDesc]);
        arrayToSheets.push([np.projectBenefits]);
        arrayToSheets.push([np.projectPlanID]);
        if (np.projectAttachmets.length != 0)
            fillupAttachments_(plan, np.projectAttachmets);
        plan.getActiveSheet().getRange('B2:B14').setValues(arrayToSheets);
        plan.getActiveSheet().getRange('B16').setValue(np.projectLocalITManager);
        saveOrUpdateProjectOnReport(np, np.saveTime, plan.getUrl(), np.requesterEmail);
    } catch (e) {
        Logger.log(e);
        sendLog(e, np.projectID);
        return {
            np: np,
            error: e.message + ' (' + e.lineNumber + ')'
        };
    }
    return {
        np: np,
        error: ''
    };
}
/*
Current status: DRAFT
*/
function sendNewRequestToApproval(np) {
    var plan;
    try {
        saveProjectNewRequest(np);
        if (np.projectID == 0) {
            np.projectID = getNewProjectID();
        }
        plan = SpreadsheetApp.openById(np.projectPlanID);
        var currentName = plan.getName();
        var newName = currentName.replace('DRAFT', np.projectID);
        plan.rename(newName);
        var comment = {
            comment: 'Request has been sent to Local Area Manager Approval.',
            date: np.saveTime,
            commenter: Session.getActiveUser().getEmail()
        };
        var sheet = plan.getActiveSheet();
        sheet.getRange(STATUS_CELL).setValue(np.projectStatus);
        sheet.getRange(ID_CELL).setValue(np.projectID);
        sheet.getRange('B15').clearContent()
        sheet.getRange('B20').clearContent();
        np.comments.push(comment);
        fillupComments_(plan, np.comments);
        //--------------------------------------
        sendEmail(np.requesterManager, 'A new request has been created and needs your attention #' + np.projectID, np.projectPlanID, np.projectName, np.requesterEmail, '');
        saveOrUpdateProjectOnReport(np, comment.date, plan.getUrl(), np.requesterManager);
        return {
            np: np,
            error: ''
        };
    } catch (e) {
        Logger.log(e);
        sendLog(e, np.projectID);
        return {
            np: np,
            error: e.message + ' (' + e.lineNumber + ')'
        };
    }
}
/*
Current status: BUSINESS_APPROVAL
*/
function saveApprovalStep(np) {
    try {
      //=============================================================================================
      if (np.projectLocalITManager == ''){
        var ITBusinessDevManager = getConfigByKey('StandardITBusinessDevelopmentManager');
        if (ITBusinessDevManager == undefined) ITBusinessDevManager = 'luciano.borges@embraco.com';
        np.projectLocalITManager = ITBusinessDevManager;
      }
      //=============================================================================================
      plan = SpreadsheetApp.openById(np.projectPlanID);
      var sheet = plan.getActiveSheet()
      var values = [];
      values.push([np.projectApproval]);
      values.push([np.projectLocalITManager.toLowerCase()]);
      values.push([np.projectApprovalComment]);
      values.push([np.projectApprovalDate]);
      values.push([np.projectApprovalApprover.toLowerCase()]);
      sheet.getRange('B15:B19').setValues(values);
      sheet.getRange('B20').clearContent();
      sheet.getRange(STATUS_CELL).setValue(np.projectStatus);
      var text = 'REJECTED';
      var comment = '';
      var currentResponsible = '';
      if (np.projectApproval) {
        text = '';
        comment = {
          comment: 'Request has been APPROVED by Local Area Manager and sent to IT Project Prioritization Forum.<br><b>Local Area Manager comment:</b> ' + np.projectApprovalComment,
          date: np.projectApprovalDate,
          commenter: np.projectApprovalApprover
        };
        sendEmail(np.projectLocalITManager, 'Request has been approved by Manager #' + np.projectID, np.projectPlanID, np.projectName, np.projectApprovalApprover, np.projectApprovalComment);
        //=========================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        var bkpITBusinessDevManager = getConfigByKey('BackupITBusinessDevelopmentManager');
        if (bkpITBusinessDevManager == undefined) bkpITBusinessDevManager = 'thalita.m.begliomini@embraco.com';
        if (np.projectLocalITManager != bkpITBusinessDevManager)
          sendEmail2ITBusinessDevManagerBKP(bkpITBusinessDevManager, 'Request has been sent IT Project Prioritization Forum #' + np.projectID, np.projectPlanID, np.projectName, np.projectBrmApprover);
        //=========================+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        currentResponsible = np.projectLocalITManager;
      } else {
        comment = {
          comment: 'Request has been REJECTED by Local Area Manager and sent back to Requester. <br><b>Local Area Manager comment: ' + np.projectApprovalComment,
          date: np.projectApprovalDate,
          commenter: np.projectApprovalApprover
        };
        sendEmail(np.requesterEmail, 'Your request has been rejected by Manager', np.projectPlanID, np.projectName, np.projectApprovalApprover, np.projectApprovalComment);
        currentResponsible = np.requesterEmail;
        np.projectStatus = projectStatus.DRAFT;
      }
      np.comments.push(comment);
      fillupComments_(plan, np.comments);
      saveOrUpdateProjectOnReport(np, comment.date, plan.getUrl(), currentResponsible);
      return {
        np: np,
        error: ''
      };
    } catch (e) {
      Logger.log(e);
      sendLog(e, np.projectID);
      return {
        np: np,
        error: e.message + ' (' + e.lineNumber + ')'
      };
    }
}
/*
Current status: IT_APPROVAL
*/
function saveITManagerApproval(np) {
    try {
        plan = SpreadsheetApp.openById(np.projectPlanID);
        var sheet = plan.getActiveSheet()
        var values = [];
        values.push([np.projectITApproval]);
        values.push([np.projectBrm.toLowerCase()]);
        values.push([np.projectITApprovalComment]);
        values.push([np.projectITApprovalDate]);
        values.push([np.projectITApprover.toLowerCase()]);
        sheet.getRange('B20:B24').setValues(values);
        sheet.getRange('D16').clearContent();
        sheet.getRange(STATUS_CELL).setValue(np.projectStatus);
        var text = '';
        var comment = '';
        var currentResponsible = '';
        if (np.projectITApproval) {
            text = '';
            comment = {
                comment: 'Request has been APPROVED by IT Project Prioritization Forum and sent to IT Manager Approval<br><b>IT Project Prioritization Forum comment:</b> ' + np.projectITApprovalComment,
                date: np.projectITApprovalDate,
                commenter: np.projectITApprover
            };
            sendEmail(np.projectBrm, 'Request has been approved by IT Project Prioritization Forum #' + np.projectID, np.projectPlanID, np.projectName, np.projectITApprover, np.projectITApprovalComment);
            currentResponsible = np.projectBrm;
        } else {
            comment = {
                comment: 'Request has been REJECTED by IT Project Prioritization Forum and sent back to Requester.<br><b>IT Project Prioritization Forum comment:</b> ' + np.projectITApprovalComment,
                date: np.projectITApprovalDate,
                commenter: np.projectITApprover
            };
            sendEmail(np.requesterEmail, 'Request has been rejected by IT Project Prioritization Forum #' + np.projectID, np.projectPlanID, np.projectName, np.projectITApprover, np.projectITApprovalComment);
            currentResponsible = np.requesterEmail;
            np.projectStatus = projectStatus.DRAFT;
        }
        np.comments.push(comment);
        fillupComments_(plan, np.comments);

        saveOrUpdateProjectOnReport(np, comment.date, plan.getUrl(), currentResponsible);
        return {
            np: np,
            error: ''
        };
    } catch (e) {
        Logger.log(e);
        sendLog(e, np.projectID);
        return {
            np: np,
            error: e.message + ' (' + e.lineNumber + ')'
        };
    }
}
/*
Current status: BRM_MANAGERS
*/
function saveBrmManagerApproval(np) {
    //
    var plan;
    try {
        plan = SpreadsheetApp.openById(np.projectPlanID);
        var sheet = plan.getActiveSheet()
        var values = [];
        values.push([np.projectBrmManagerApproval]);
        values.push([np.projectBrmManagerBrmAnalyst]);
        values.push([np.projectBrmManagerApprovalComment]);
        values.push([np.projectBrmManagerApprovalDate]);
        values.push([np.projectBrmManagerApprover]);
        sheet.getRange('D16:D20').setValues(values);
        sheet.getRange('B25').clearContent();
        sheet.getRange(STATUS_CELL).setValue(np.projectStatus);
        //=====================================
        var text = '';
        var comment = '';
        var currentResponsible = '';
        if (np.projectBrmManagerApproval) {
            text = '';
            comment = {
                comment: 'Request has been APPROVED by IT Manager Approval and sent to BRM Analyst.<br><b>BRM comment:</b> ' + np.projectBrmManagerApprovalComment,
                date: np.projectBrmManagerApprovalDate,
                commenter: np.projectBrmManagerApprover
            };
            sendEmail(np.projectBrmManagerBrmAnalyst, 'Request has been approved by IT Manager Approval #' + np.projectID, np.projectPlanID, np.projectName, np.projectBrmManagerApprover, np.projectBrmManagerApprovalComment);
            currentResponsible = np.projectBrmManagerBrmAnalyst;
            np.projectStatus = projectStatus.PROJECT_ANALYSIS
        } else {
            // TODO: VERIFICAR PRA QUEM VOLTA
            comment = {
                comment: 'Request has been REJECTED by IT Manager Approval and sent back to Requester.<br><b>IT Manager Approval comment:</b> ' + np.projectBrmManagerApprovalComment,
                date: np.projectBrmManagerApprovalDate,
                commenter: np.projectBrmManagerApprover
            };
            sendEmail(np.requesterEmail, 'Request has been rejected by IT Manager Approval #' + np.projectID, np.projectPlanID, np.projectName, np.projectBrmManagerApprover, np.projectBrmManagerApprovalComment);
            currentResponsible = np.requesterEmail;
            np.projectStatus = projectStatus.DRAFT;
        }
        np.comments.push(comment);
        fillupComments_(plan, np.comments);
        saveOrUpdateProjectOnReport(np, comment.date, plan.getUrl(), currentResponsible);
        return {
            np: np,
            error: ''
        };
    } catch (e) {
        Logger.log(e);
        sendLog(e, np.projectID);
        return {
            np: np,
            error: e.message + ' (' + e.lineNumber + ')'
        };
    }


}
/*
Current status: PROJECT_ANALYSIS
*/
function saveBrmAnalystApproval(np) {
    var plan;
    try {
        plan = SpreadsheetApp.openById(np.projectPlanID);
        var sheet = plan.getActiveSheet()
        var values = [];
        values.push([np.projectBrmApproval]);
        values.push([np.projectBrmPortfolioManager.toLowerCase()]);
        values.push([np.projectBrmApprovalComment]);
        values.push([np.projectBrmApprovalDate]);
        values.push([np.projectBrmApprover.toLowerCase()]);
        sheet.getRange('B25:B29').setValues(values);
        sheet.getRange('B30').clearContent();
        sheet.getRange(STATUS_CELL).setValue(np.projectStatus);
        var text = '';
        var comment = '';
        var currentResponsible = '';
        if (np.projectBrmApproval) {
            text = '';
            comment = {
                comment: 'Request has been APPROVED by BRM Analyst and sent to IT Portfolio Manager.<br><b>BRM Analyst comment:</b> ' + np.projectBrmApprovalComment,
                date: np.projectBrmApprovalDate,
                commenter: np.projectBrmApprover
            };
            sendEmail(np.projectBrmPortfolioManager, 'Request has been approved by BRM Analyst #' + np.projectID, np.projectPlanID, np.projectName, np.projectBrmApprover, np.projectBrmApprovalComment);
            currentResponsible = np.projectBrmPortfolioManager;
            var standardPortfolioManager = getConfigByKey('StandardPorfolioManager');
            if (standardPortfolioManager == undefined) standardPorfolioManager = 'thalita.m.begliomini@embraco.com';
            if (np.projectBrmPortfolioManager != standardPortfolioManager)
                sendEmail2Portfolio(standardPortfolioManager, 'Request has been sent to Portfolio Manager #' + np.projectID, np.projectPlanID, np.projectName, np.projectBrmApprover, np.projectBrmPortfolioManager);
        } else {
            comment = {
                comment: 'Request has been REJECTED by BRM Analyst and sent back to BRM.<br><b>BRM Analyst comment:</b> ' + np.projectBrmApprovalComment,
                date: np.projectBrmApprovalDate,
                commenter: np.projectBrmApprover
            };
            sendEmail(np.projectLocalITManager, 'Request has been rejected by BRM Analyst #' + np.projectID, np.projectPlanID, np.projectName, np.projectBrmApprover, np.projectBrmApprovalComment);
            currentResponsible = np.projectLocalITManager;
            np.projectStatus = projectStatus.BRM_MANAGERS;
        }
        np.comments.push(comment);
        fillupComments_(plan, np.comments);
        saveOrUpdateProjectOnReport(np, comment.date, plan.getUrl(), currentResponsible);
        return {
            np: np,
            error: ''
        };
    } catch (e) {
        Logger.log(e);
        sendLog(e, np.projectID);
        return {
            np: np,
            error: e.message + ' (' + e.lineNumber + ')'
        };
    }
}
/*
Current status: PROJECT_PRIORITIZATION_FOR_EXECUTION
*/
function savePortfolioApproval(np) {
    var plan;
    try {
        plan = SpreadsheetApp.openById(np.projectPlanID);
        var sheet = plan.getActiveSheet()
        var values = [];
        values.push([np.projectITPortfolioApproval]);
        values.push([np.projectPm.toLowerCase()]);
        values.push([np.projectITPortfolioComment]);
        values.push([np.projectITPortfolioApprovalDate]);
        values.push([np.projectITPortfolioApprover.toLowerCase()]);
        sheet.getRange('B30:B34').setValues(values);
        sheet.getRange(STATUS_CELL).setValue(np.projectStatus);
        var text = '';
        var comment = '';
        var currentResponsible = '';
        if (np.projectITPortfolioApproval) {
            text = '';
            comment = {
                comment: 'Request has been APPROVED by IT Portfolio Manager and sent to Project Manager.<br><b>IT Portfolio Manager comment:</b> ' + np.projectITPortfolioComment,
                date: np.projectITPortfolioApprovalDate,
                commenter: np.projectITPortfolioApprover
            };
            sendEmail(np.projectPm, 'Request has been approved by IT Portfolio Manager #' + np.projectID, np.projectPlanID, np.projectName, np.projectITPortfolioApprover, np.projectITPortfolioComment);
            currentResponsible = np.projectPm;
        } else {
            comment = {
                comment: 'Request has been REJECTED by IT Portfolio Manager and sent back to BRM Analyst.<br><b>IT Portfolio Manager comment:</b> ' + np.projectITPortfolioComment,
                date: np.projectITPortfolioApprovalDate,
                commenter: np.projectITPortfolioApprover
            };
            sendEmail(np.projectBrm, 'Request has been rejected by IT Portfolio Manager #' + np.projectID, np.projectPlanID, np.projectName, np.projectITPortfolioApprover, np.projectITPortfolioComment);
            currentResponsible = np.projectBrm;
            np.projectStatus = projectStatus.PROJECT_ANALYSIS;
        }
        np.comments.push(comment);
        fillupComments_(plan, np.comments);
        saveOrUpdateProjectOnReport(np, comment.date, plan.getUrl(), currentResponsible);
        //---------------------------------------
        sendEmail2ITServiceLeader(np.projectPlanID, np.projectName, np.projectITPortfolioApprover, np.projectID, np.requesterLocation);
        return {
            np: np,
            error: ''
        };
    } catch (e) {
        Logger.log(e);
        sendLog(e, np.projectID);
        return {
            np: np,
            error: e.message + ' (' + e.lineNumber + ')'
        };
    }
}
/*
    Current status: PROJECT_EXECUTION
*/
function saveProjectManagerRequest(np) {
    var plan;
    try {
        plan = SpreadsheetApp.openById(np.projectPlanID);
        var sheet = plan.getActiveSheet()
        var values = [];
        values.push([np.projectSponsorName]);
        values.push([np.projectSponsorEmail.toLowerCase()]);
        values.push([np.projectPmApprovalDate]);
        values.push([np.projectPmApprover.toLowerCase()]);
        values.push([np.projectPmComment]);
        sheet.getRange('B35:B39').setValues(values);
        sheet.getRange(STATUS_CELL).setValue(np.projectStatus);
        var comment = '';
        var nextResponsible = '';
        var addToAutoClose = false;
        if (np.projectStatus == projectStatus.PROJECT_IN_EVALUATION) {
            comment = {
                comment: 'Request has been sent to Sponsor for evaluation.',
                date: np.projectPmApprovalDate,
                commenter: np.projectPmApprover
            };
            sendEmail(np.projectSponsorEmail, 'Request needs your attention #' + np.projectID, np.projectPlanID, np.projectName, np.projectPmApprover, '');
            nextResponsible = np.projectSponsorEmail;
            addToAutoClose = true;
        } else {
            comment = {
                comment: 'Request has been saved with a new comment by Project Manager.<br><b>Project Manager comment:</b> ' + np.projectPmComment,
                date: np.projectPmApprovalDate,
                commenter: np.projectPmApprover
            };
            nextResponsible = np.projectPmApprover;
            np.projectStatus = projectStatus.PROJECT_PRIORITIZATION_FOR_EXECUTION;
        }
        np.comments.push(comment);
        fillupComments_(plan, np.comments);
        saveOrUpdateProjectOnReport(np, np.projectPmApprovalDate, plan.getUrl(), nextResponsible);
        if (addToAutoClose == true) addRequestToAutoClose(np);
        return {
            np: np,
            error: ''
        };
    } catch (e) {
        Logger.log(e);
        sendLog(e, np.projectID);
        return {
            np: np,
            error: e.message + ' (' + e.lineNumber + ')'
        };
    }
}
/*
Current status: PROJECT_IN_EVALUATION
*/
function saveEvaluationForm(np) {
    var plan;
    try {
        plan = SpreadsheetApp.openById(np.projectPlanID);
        var sheet = plan.getActiveSheet()
        var values = [];
        np.answerAuthor = Session.getActiveUser().getEmail();
        values.push([np.question1]);
        values.push([np.question2]);
        values.push([np.question3]);
        values.push([np.question4]);
        values.push([np.question5]);
        values.push([np.answerDate]);
        values.push([np.answerAuthor]);
        sheet.getRange('D9:D15').setValues(values);
        sheet.getRange(STATUS_CELL).setValue(np.projectStatus);
        var comment = {
            comment: 'Evaluation has been answer and request finished.',
            date: np.answerDate,
            commenter: Session.getActiveUser().getEmail()
        };
        sendEmail(np.projectPm, 'Evaluation has been answer and request finished #' + np.projectID, np.projectPlanID, np.projectName, Session.getActiveUser().getEmail(), '');
        np.comments.push(comment);
        fillupComments_(plan, np.comments);
        saveOrUpdateProjectOnReport(np, comment.date, plan.getUrl(), np.projectPm);
        //-------------------------------------------
        removeRequestFromAutoClose(np.projectID);
        return {
            np: np,
            error: ''
        };
    } catch (e) {
        Logger.log(e);
        sendLog(e, np.projectID);
        return {
            np: np,
            error: e.message + ' (' + e.lineNumber + ')'
        };
    }
}
//----------------------------------------------------------------------------
function cancelRequest(np) {
    var plan;
    try {
        plan = SpreadsheetApp.openById(np.projectPlanID);
        var sheet = plan.getActiveSheet()
        var values = [];
        values.push([np.projectCanceled]);
        values.push([np.projectCanceledDate]);
        values.push([np.projectCanceledAuthor]);
        values.push([np.projectCanceledAuthorComment]);
        sheet.getRange('D2:D5').setValues(values);
        if (np.projectCanceledRequesterAnswer == true) {
            sheet.getRange('D6').setValue(np.projectCanceledRequesterAnswer);
        }
        var comment = {
            comment: 'Request has been CANCELED by ' + np.projectCanceledAuthor + '.<br><b>Cancelation comment:</b> ' + np.projectCanceledAuthorComment,
            date: np.projectCanceledDate,
            commenter: np.projectCanceledAuthor
        };
        sendEmail(np.projectCanceledAuthor, 'Request has been CANCELED #' + np.projectID, np.projectPlanID, np.projectName, np.projectCanceledAuthor, np.projectApprovalComment);
        np.comments.push(comment);
        fillupComments_(plan, np.comments);
        var tmpStatus = np.projectStatus;
        np.projectStatus = projectStatus.CANCELED;
        saveOrUpdateProjectOnReport(np, comment.date, plan.getUrl(), np.projectCanceledAuthor);
        np.projectStatus = tmpStatus;
        return {
            np: np,
            error: ''
        };
    } catch (e) {
        Logger.log(e);
        sendLog(e, np.projectID);
        return {
            np: np,
            error: e.message + ' (' + e.lineNumber + ')'
        };
    }
}

function updateCancelRequest(np) {
    var plan;
    try {
        plan = SpreadsheetApp.openById(np.projectPlanID);
        var sheet = plan.getActiveSheet()
        var values = [];
        values.push([np.projectCanceledRequesterAnswer]);
        values.push([np.projectCanceledRequesterAnswerDate]);
        values.push([np.projectCanceledRequesterAnswerComment]);
        sheet.getRange('D6:D8').setValues(values);
        var text = 'Request has been confimed as CANCELED.<br><b>Requester comment:</b> ' + np.projectCanceledRequesterAnswerComment;
        if (np.projectCanceledRequesterAnswer === false) {
            text = 'Request has been RESTORED.<br><b>Requester comment:</b> ' + np.projectCanceledRequesterAnswerComment;
        }
        var comment = {
            comment: text,
            date: np.projectCanceledRequesterAnswerDate,
            commenter: np.requesterEmail
        };
        sendEmail(np.projectPm, 'Request has been RESTORED #' + np.projectID, np.projectPlanID, np.projectName, np.requesterEmail, np.projectCanceledRequesterAnswerComment);
        np.comments.push(comment);
        fillupComments_(plan, np.comments);
        sheet.getRange('D2').setValue(np.projectCanceledRequesterAnswer);
        if (np.projectCanceledRequesterAnswer == false) {
            saveOrUpdateProjectOnReport(np, np.projectCanceledRequesterAnswerDate, plan.getUrl(), np.requesterEmail);
        }
        /*else {
          np.projectStatus = projectStatus.CANCELED;
          sheet.getRange(STATUS_CELL).setValue(np.projectStatus);
        }*/
        return {
            np: np,
            error: ''
        };
    } catch (e) {
        Logger.log(e);
        sendLog(e, np.projectID);
        return {
            np: np,
            error: e.message + ' (' + e.lineNumber + ')'
        };
    }
}

function cancelDraft(np) {
    var plan;
    var tmpProjectPlanID = np.projectPlanID;
    try {
        plan = DriveApp.getFileById(np.projectPlanID);
        plan.setTrashed(true);
        np = newProject();

        var report = SpreadsheetApp.openById(DB).getSheetByName('Requests');
        var rows = report.getRange('A6:A').getFormulas();
        var regexp = /id=(.*)\"\,/
        for (var i in rows) {
            var row = rows[i];
            if (row == '') continue;
            var result = regexp.exec(row[0]);
            var idd = result[1];
            if (idd === tmpProjectPlanID) {
                report.deleteRow((+i + 6));
                break;
            }
        }
        return {
            np: np,
            error: ''
        };
    } catch (e) {
        Logger.log(e);
        sendLog(e, np.projectID);
        return {
            np: np,
            error: e.message + ' (' + e.lineNumber + ')'
        };
    }
}
//=============================================================================================================
function retrieveProject(id) {
    try {
        var np = {};
        var file = DriveApp.getFileById(id);
        if (file.isTrashed()) {
            throw new ProjectException('Project not found', 475);
        }
        var plan = SpreadsheetApp.openById(id);
        var values = plan.getRangeByName('ProjectData').getValues();
        for (var i in values) {
            var line = values[i]
            np[line[0]] = line[1];
        }
        var values = plan.getRangeByName('ProjectData2').getValues();
        for (var i in values) {
            var line = values[i]
            np[line[0]] = line[1];
        }
        var comments = [];
        var commentsValues = plan.getSheetByName('RequestComments').getRange(COMMENT_RANGE).getValues();
        for (var i in commentsValues) {
            if (commentsValues[i][0] == "") break;
            comments.push({
                comment: commentsValues[i][0],
                commenter: commentsValues[i][1],
                date: commentsValues[i][2]
            });
        }
        np.comments = comments;
        var attachmets = [];
        var attachmetsValues = plan.getSheetByName('RequestAttachments').getRange(ATTACHMENTS_RANGE).getValues();
        for (var i in attachmetsValues) {
            if (attachmetsValues[i][0] == "") break;
            attachmets.push({
                name: attachmetsValues[i][0],
                date: attachmetsValues[i][1],
                url: attachmetsValues[i][2],
                icon: attachmetsValues[i][3],
                userEmail: attachmetsValues[i][4],
                id: attachmetsValues[i][5]
            });
        }

        np.projectAttachmets = attachmets;
        np.url = plan.getUrl();
        return {
            np: np,
            error: ''
        };
    } catch (e) {
        Logger.log(e);
        sendLog(e, np.projectID);
        return {
            np: np,
            error: e.message + ' (' + e.lineNumber + ')'
        };
    }
}
