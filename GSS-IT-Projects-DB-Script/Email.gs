function sendEmail(emailTemplate, recipient, subject, planId, prjName, author, comment){
  try{
    var body = HtmlService.createHtmlOutputFromFile(emailTemplate).getContent();
    var logo = DriveApp.getFileById('0B973TWs84ZCEVW9PcW5qY0lzY3M').getBlob().setName('embracoLogo');
    body = body.replace('$PLAN_ID$', planId);
    body = body.replace('$PRJ_NAME$', prjName);
    body = body.replace('$AUTHOR$', author);
    body = body.replace('$URL$', this.APP_URL);

    if (comment == '') comment = 'None.';
    body = body.replace('$COMMENT$', comment);
    var options = {
      htmlBody:body,
      inlineImages: { embracoLogo: logo },
      name:'IT Project Request'
    };
    GmailApp.sendEmail(recipient, '[IT Project Request] '+subject, '', options);
  } catch(e){
    SpreadsheetApp.openById('18i2v2Ili1AgRs2Ei2Pi--NmO8VrvIZc9z5-VLGKTyAI').getActiveSheet().appendRow([Session.getActiveUser().getEmail(), new Date().toString(), e]);
  }
}
