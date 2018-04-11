function doGet(e) {
    var html = HtmlService.createTemplateFromFile('index');

    if (e.parameter.id) {
        var id = e.parameter.id;
        if (id.substring(0, 7) == 'embraco') id = id.replace('embraco', '');
        html.id = id;
    } else {
        html.id = '0';
    }
    var url = ScriptApp.getService().getUrl();
    html.url = url;

    return html.evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setTitle("Embraco - IT Project Request")
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, minimum-scale=1.0');
}

function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getOAuthToken() {
    var root = DriveApp.getRootFolder();
    var token = ScriptApp.getOAuthToken();
    var obj = {
        root: root,
        token: token
    };
    return obj;
}
//*************************************************************************************
//*************************************************************************************
//*************************************************************************************
function searchUser(nome) {
    Logger.log(nome);
    var usersList = [];
    var args = {
        domain: 'embraco.com',
        query: nome,
        fields: 'users(name/fullName,primaryEmail)',
        maxResults: 30
    }
    var page = AdminDirectory.Users.list(args);
    var users = page.users;
    if (users) {
        for (var i = 0; i < users.length; i++) {
            var user = users[i];
            usersList.push(user.name.fullName + ' - ' + user.primaryEmail);
        }
    } else {
        usersList.push('No users found.');
    }
    return usersList;
}

function sendEmail(recipient, subject, planId, prjName, author, comment) {
    try {
        var body = HtmlService.createHtmlOutputFromFile('email_content').getContent();
        var logo = DriveApp.getFileById('0B973TWs84ZCEVW9PcW5qY0lzY3M').getBlob().setName('embracoLogo');
        body = body.replace('$PLAN_ID$', planId);
        body = body.replace('$PRJ_NAME$', prjName);
        body = body.replace('$AUTHOR$', author);
        body = body.replace('$URL$', this.APP_URL);

        if (comment == '') comment = 'None.';
        body = body.replace('$COMMENT$', comment);
        var options = {
            htmlBody: body,
            inlineImages: {
                embracoLogo: logo
            },
            name: 'IT Project Request'
        };
        GmailApp.sendEmail(recipient, '[IT Project Request] ' + subject, '', options);
    } catch (e) {
        sendLog(e, prjName);
    }
}

function sendEmail2ITServiceLeader(planId, prjName, author, id, location) {
    try {
        var subject = 'Notification';
        var body = HtmlService.createHtmlOutputFromFile('emailITServiceLeader').getContent();
        var logo = DriveApp.getFileById('0B973TWs84ZCEVW9PcW5qY0lzY3M').getBlob().setName('embracoLogo');
        body = body.replace('$PLAN_ID$', planId);
        body = body.replace('$PRJ_NAME$', prjName);
        body = body.replace('$AUTHOR$', author);
        body = body.replace('$URL$', this.APP_URL);

        //ITServiceNotifictionEmail-BRAZIL copiar do HW
        var configs = SpreadsheetApp.openById(DB).getSheetByName('Configs').getRange('A3:B').getValues();
        var itemsDB = configs.filter(function(current) {
            if (current[0].indexOf('ITServiceNotificationEmail') > -1)
                return current;
        });
        var recipient = '';
        for (var i in itemsDB) {
            var current = itemsDB[i];
            var name = current[0].split('-')[1];
            if (name.toUpperCase() == location.toUpperCase()) {
                recipient = current[1];
                break;
            }
        }
        if (recipient == '') return true;

        var comment = 'None.';
        body = body.replace('$COMMENT$', comment);
        var options = {
            htmlBody: body,
            inlineImages: {
                embracoLogo: logo
            },
            name: 'IT Project Request'
        };
        //    recipient = 'it.apps@embraco.com';
        GmailApp.sendEmail(recipient, '[IT Project Request] ' + subject + ' #' + id, '', options);
    } catch (e) {
        sendLog(e, prjName);
    }
}

function sendEmail2Portfolio(recipient, subject, planId, prjName, author, portfolioemail) {
    try {
        var body = HtmlService.createHtmlOutputFromFile('emailPortfolio').getContent();
        var logo = DriveApp.getFileById('0B973TWs84ZCEVW9PcW5qY0lzY3M').getBlob().setName('embracoLogo');
        body = body.replace('$PLAN_ID$', planId);
        body = body.replace('$PRJ_NAME$', prjName);
        body = body.replace('$AUTHOR$', author);
        body = body.replace('$URL$', this.APP_URL);
        body = body.replace('$PORTFOLIO$', portfolioemail);

        var options = {
            htmlBody: body,
            inlineImages: {
                embracoLogo: logo
            },
            name: 'IT Project Request'
        };
        GmailApp.sendEmail(recipient, '[IT Project Request] ' + subject, '', options);
    } catch (e) {
        sendLog(e, prjName);
    }
}

function sendEmail2ITBusinessDevManagerBKP(recipient, subject, planId, prjName, author) {
    try {
        var body = HtmlService.createHtmlOutputFromFile('emailITBusinessDevManagerBKP').getContent();
        var logo = DriveApp.getFileById('0B973TWs84ZCEVW9PcW5qY0lzY3M').getBlob().setName('embracoLogo');
        body = body.replace('$PLAN_ID$', planId);
        body = body.replace('$PRJ_NAME$', prjName);
        body = body.replace('$AUTHOR$', author);
        body = body.replace('$URL$', this.APP_URL);

        var options = {
            htmlBody: body,
            inlineImages: {
                embracoLogo: logo
            },
            name: 'IT Project Request'
        };
        GmailApp.sendEmail(recipient, '[IT Project Request] ' + subject, '', options);
    } catch (e) {
        sendLog(e, prjName);
    }
}

function getStatusString(status) {
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
            return 'IT PROJECT LEADER ANALYSIS';
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

function locationPlants_(location) {
    switch (location) {
        case 'BRAZIL PLANT':
            return 'BRAZIL';
        case 'CORPORATE':
            return 'BRAZIL';
        case 'CHINA BEIJING PLANT':
            return 'BEIJING CHINA';
        case 'EE QINGDAO':
            return 'QD CHINA';
        case 'ITALY PLANT':
            return 'ITALY';
        case 'MEXICO PLANT':
            return 'MEXICO';
        case 'ENA OFFICE':
            return 'usa';
        case 'RUSSIA OFFICE':
            return 'RUSSIA';
        case 'SLOVAKIA PLANT':
            return 'SLOVAKIA';
        default:
            return '';
    }
}

function retrieveCombosData() {
    var db = SpreadsheetApp.openById(DB).getSheetByName('Combos Data');

    var values = db.getRange('A3:D').getValues();
    var combos = {};
    var localitmanager = [{
        display: '',
        value: ''
    }];
    var portfolio = [];
    for (var i in values) {
        var line = values[i];
        if (line[0] != '') localitmanager.push({
            display: line[0],
            value: line[1]
        });
        if (line[2] != '') portfolio.push({
            display: line[2],
            value: line[3]
        });
    }
    combos.localitmanager = localitmanager;
    combos.portfolio = portfolio;
    return combos;
}

function fillupAttachments_(plan, files) {
    plan.getSheetByName('RequestAttachments').getRange(ATTACHMENTS_RANGE).clearContent();
    var arrayToSheets = [];
    for (var i in files) {
        arrayToSheets.push([files[i].name, files[i].date, files[i].url, files[i].icon, files[i].userEmail, files[i].id]);
    }
    if (arrayToSheets.length > 0)
        plan.getSheetByName('RequestAttachments').getRange(ATTACHMENTS_RANGE + (Number(arrayToSheets.length) + 2)).setValues(arrayToSheets);
}

function fillupComments_(plan, comments) {
    plan.getSheetByName('RequestComments').getRange(COMMENT_RANGE).clearContent();
    var arrayToSheets = [];
    for (var i in comments) {
        arrayToSheets.push([comments[i].comment, comments[i].commenter, comments[i].date]);
    }
    plan.getSheetByName('RequestComments').getRange(COMMENT_RANGE + (Number(arrayToSheets.length) + 2)).setValues(arrayToSheets);
}

function getNewProjectID() {
    var lock = LockService.getScriptLock();
    lock.waitLock(2000);
    var db = SpreadsheetApp.openById(DB).getSheetByName('Configs');
    var lastIDCell = db.getRange('B2');
    var lastIDValue = Number(lastIDCell.getValue()) + 1;
    lastIDCell.setValue(lastIDValue);
    Utilities.sleep(1500);
    lock.releaseLock();
    return lastIDValue;
}

function saveFiles(np) {
    try {
        var plan = SpreadsheetApp.openById(np.projectPlanID);
        var comment = {
            comment: 'Files on the Request has been updated.',
            date: np.saveTime,
            commenter: Session.getActiveUser().getEmail()
        };
        np.comments.push(comment);
        fillupComments_(plan, np.comments);
        fillupAttachments_(plan, np.projectAttachmets);
    } catch (e) {
        Logger.log(e);
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

function ProjectException(message, line) {
    this.message = message;
    this.name = "ProjectException";
    this.lineNumber = line;
}

function usersCache() {
    var plan = SpreadsheetApp.openById(USER_DBCACHE);
    var sheet = plan.getSheetByName('Users');
    var users = [];
    var values = sheet.getRange('A2:B').getValues();
    for (var i in values) {
        users.push({
            text: values[i][0],
            value: values[i][1]
        });
    }
    return users;
}

function retrieveUserInfo(email, np) {
    var optional = {
        fields: 'name/fullName,organizations,phones,relations'
    }
    var detail = AdminDirectory.Users.get(email, optional);
    np.requesterName = detail.name.fullName;
    np.requesterEmail = email;
    var phone = detail.phones;
    if (phone != undefined) {
        var p = phone.filter(function(it) {
            return it.type == 'work'
        });
        np.requesterPhone = p[0].value;
    }
    var manager = detail.relations;
    if (manager != undefined) {
        var m = manager.filter(function(it) {
            return it.type == 'manager'
        });
        np.requesterManager = m[0].value;
    }
    if (detail.organizations != undefined) {
        np.requesterDepartment = detail.organizations[0].department;
        np.requesterLocation = locationPlants_(detail.organizations[0].name);
    }
    return np;

}

function updateall_() {
    var re = /id=(.*)", "/;
    var user = Session.getActiveUser().getEmail();
    var plan = SpreadsheetApp.openById('17JKCj7gFAC4veKp7glbO3jtLwLCNbql7URzTAo8z1l4').getSheetByName('Requests');
    var allData = plan.getRange('A6:J').getDisplayValues();
    var rawData = plan.getRange('A6:A').getFormulas();
    var data = [];
    for (var i in allData) {
        var row = allData[i];
        if (row[0] == '') continue;
        var status = row[1];
        var requester = row[4];
        var manager = row[5];
        var col1 = rawData[i][0];
        var id = re.exec(col1)[1];
        var np = retrieveProject(id).np;
        saveOrUpdateProjectOnReport(np, row[9], np.url, row[2]);
        Logger.log(np.projectID);
    }
    Logger.log(data);
    //  saveOrUpdateProjectOnReport();

}
//-------------------------------------------------------------------------------------------------------
//-------------------------------------------------------------------------------------------------------
function saveOrUpdateProjectOnReport(np, lastUpdate, url, currentResponsible) {
    var report = SpreadsheetApp.openById(DB).getSheetByName('Requests');
    var rows = report.getRange('A6:A').getFormulas();
    var found = false;
    var regexp = /id=(.*)\"\,/
    for (var i in rows) {
        var row = rows[i];
        if (row == '') continue;
        var result = regexp.exec(row[0]);
        var idd = result[1];
        if (idd == np.projectPlanID) {
            var text = np.projectID;
            if (np.projectID == 0) text = 'DRAFT';
            var questions = {
                question1: np.question1,
                question2: np.question2,
                question3: np.question3,
                question4: np.question4,
                question5: np.question5,
                answerDate: np.answerDate,
                answerAuthor: np.answerAuthor
            };
            var values = ['=HYPERLINK("' + APP_URL + '?id=' + np.projectPlanID + '", "' + text + '")',
                getStatusString(np.projectStatus),
                currentResponsible,
                new Date(np.requestDate),
                np.requesterEmail,
                np.requesterManager,
                '=HYPERLINK("' + url + '", "' + np.projectName + '")',
                np.requesterDepartment,
                np.requesterLocation,
                new Date(lastUpdate),
                getString(np.projectApproval),
                np.projectLocalITManager,
                getString(np.projectITApproval),
                np.projectBrm,
                getString(np.projectBrmManagerApproval),
                np.projectBrmManagerBrmAnalyst,
                getString(np.projectBrmApproval),
                np.projectBrmPortfolioManager,
                getString(np.projectITPortfolioApproval),
                np.projectPm,
                np.projectSponsorEmail,
                JSON.stringify(questions)
            ];
            report.getRange((+i + 6), 1, 1, values.length).setValues([values]);
            found = true;
            break;
        }
    }
    if (found == false) {
        // LOCK SECTION!
        var lock = LockService.getScriptLock();
        lock.waitLock(2000);
        var text = np.projectID;
        if (np.projectID == 0) text = 'DRAFT';
        report.insertRowBefore(6);
        var values = ['=HYPERLINK("' + APP_URL + '?id=' + np.projectPlanID + '", "' + text + '")',
            'DRAFT',
            currentResponsible,
            new Date(np.requestDate),
            np.requesterEmail,
            np.requesterManager,
            '=HYPERLINK("' + url + '", "' + np.projectName + '")',
            np.requesterDepartment,
            np.requesterLocation,
            new Date(lastUpdate),
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            ''
        ];
        report.getRange(6, 1, 1, values.length).setValues([values]);
        lock.releaseLock();
        // LOCK SECTION END
    }
}

function removeRequestFromAutoClose(id) {
    var report = SpreadsheetApp.openById(DB).getSheetByName('Requests to close');
    var rows = report.getRange('A2:A').getDisplayValues();
    for (var i in rows) {
        var row = rows[i];
        if (row == '') break;
        var idd = row[0];
        if (idd == id) {
            report.deleteRow(+i + Number(2));
            break;
        }
    }
}

function addRequestToAutoClose(np) {
    var plan;
    try {
        plan = SpreadsheetApp.openById(DB).getSheetByName('Requests to close');
        var DAYS_TO_CLOSE = getConfigByKey('autoCloseDays');
        if (DAYS_TO_CLOSE == undefined) DAYS_TO_CLOSE = 7;
        plan.appendRow([np.projectID, np.projectPlanID, moment().add(DAYS_TO_CLOSE, 'days').format('YYYY/MM/DD')]);
    } catch (e) {
        Logger.log(e);
        return {
            np: np,
            error: e.message + ' (' + e.lineNumber + ')'
        };
    }
}

function getConfigByKey(key) {
    var db = SpreadsheetApp.openById(DB).getSheetByName('Configs');
    var allValues = db.getRange('A2:B').getValues();
    for (var i in allValues) {
        if (allValues[i][0] == key)
            return allValues[i][1];
    }
    return undefined;
}

function getString(status) {
    if (status === true)
        return 'APPROVED';
    else {
        if (status === false)
            return 'REJECTED';
        else
            return '';
    }
}

function sendLog(e, id) {
    SpreadsheetApp.openById(LOGS_PLAN).getActiveSheet().appendRow(['IT Projects', id, Session.getActiveUser().getEmail(), new Date().toString(), e.message + ' (' + e.lineNumber + ')']);
}
