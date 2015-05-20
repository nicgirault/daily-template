function sendMail(points, draftId){
  var draft = GmailApp.getMessageById(draftId);
  var template = draft.getBody();
  var subjectTemplate = draft.getSubject();
  var result = getContentFromTemplate(template, parseInt(points), false);
  var subject = buildSubject(subjectTemplate);

  MailApp.sendEmail(draft.getTo(), subject, 'Impossible de lire le contenu', {
    cc: draft.getCc(),
    htmlBody: result[0],
    inlineImages: result[1]
 });
}


function previewEmail(formObject){
  var draft = GmailApp.getMessageById(formObject.draft);
  var bodyTemplate = draft.getBody();
  var subjectTemplate = draft.getSubject();
  var body = getContentFromTemplate(bodyTemplate, formObject.pointsToValidate, true);
  var subject = buildSubject(subjectTemplate);

  var template = HtmlService.createTemplateFromFile('previewDaily');
  template.to = draft.getTo();
  template.cc = draft.getCc();
  template.subject = subject;
  template.body = body;
  template.draftId = formObject.draft;
  template.points = formObject.pointsToValidate;

  var html = template.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(800)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Preview daily mail');
}

function openDailyMailForm() {
  var template = HtmlService.createTemplateFromFile('dailyMailForm');
  var drafts = GmailApp.getDraftMessages();

  filtered_drafts = [];
  for(var i=0;i<drafts.length;i++){
    if(drafts[i].getSubject() != ''){
      filtered_drafts.push(drafts[i]);
    }
  }
  template.drafts = filtered_drafts;
  template.exampleSubject = getSubjectExample();
  template.exampleBody = getBodyExample();
  template.emailQuotaRemaining = MailApp.getRemainingDailyQuota();

  var html = template.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Daily Mail');
}


function buildSubject(subject){
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // subject [BNP Market Tools] Daily report ##sprint## - Jour ##sprintDay## (##date##)
  var sprintNumber = ss.getRangeByName('sprintNumber').getValue();
  var startDate = ss.getRangeByName('startDate').getValue();
  startDate = moment(startDate);

  var sprintDay = startDate.diff(moment(), 'days');

  subject = subject.replace('{sprintNumber}', sprintNumber);
  subject = subject.replace('{sprintDay}', sprintDay);
  subject = subject.replace('{date}', moment().format('LL'));
  return subject;
}

function getContentFromTemplate(template, toValidatePoints, preview){
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sprintNumber = ss.getRangeByName('sprintNumber').getValue();
  var sprintGoal = ss.getRangeByName('sprintGoal').getValue();
  var doneValues = ss.getRangeByName('done').getValues();
  var standardValues = ss.getRangeByName('standard').getValues();
  var totalPoints = ss.getRangeByName('totalPoints').getValue();
  var standard = 0;
  var done = 0;
  for(var i=0;i<doneValues.length;i++){
    var value = doneValues[i][0];
    if(value != ''){
      standard = standardValues[i][0];
      done = doneValues[i][0];
    }
  }

  var donePoints = totalPoints - done;
  var toStandardPoints = standard - done;
  var earlyOrLate = '';
  var validationColor = '';
  var doneColor = '';
  if(toValidatePoints > 0){
    validationColor = '#fb8072';
  }
  else {
    validationColor = '#b3de69';
  }
  if(toStandardPoints >= 0){
    earlyOrLate = 'Avance';
    doneColor = '#b3de69';
  }
  else {
    earlyOrLate = 'Retard';
    doneColor = '#fb8072';
  }

  var html = template
    .split('{sprintNumber}').join(''+sprintNumber)
    .split('{sprintGoal}').join(''+sprintGoal)
    .split('{totalPoints}').join(''+totalPoints)
    .split('{donePoints}').join(''+donePoints)
    .split('{toValidatePoints}').join(''+toValidatePoints)
    .split('{toStandardPoints}').join(''+toStandardPoints)
    .split('{earlyOrLate}').join(earlyOrLate)
    .split('{totalPoints}').join(''+totalPoints)
    .split('{doneColorS}').join('<span style="color: '+doneColor+'">')
    .split('{doneColorE}').join('</span>')
    .split('{validationColorS}').join('<span style="color: '+validationColor+'">')
    .split('{validationColorE}').join('</span>');

  if(preview) {
    return getImages(html);
  } else {
    return fetchImages(html);
  }
}

function fetchImages(template) {
  template = template.split('{bdc}').join('<img src="cid:bdc" />')
  var inlineImages = {
    bdc: getBDC().getAs('image/png')
  };

  var html = template.replace(/IMG.*?href="(.*?)".*?IMG/g, function(whole, url) {
    url = url.replace('&amp;', '&');
    var response = UrlFetchApp.fetch(url);
    imgBlob = response.getAs('image/png');

    var id = makeid();
    inlineImages[id] = imgBlob;

    var balise = '<img src="cid:'+id+'">';
    return balise;
  });

  return [html, inlineImages];
}

function getImages(template) {
  var chart64 = Utilities.base64Encode(getBDC().getBlob().getBytes());
  var html = template
    .replace(/IMG.*?href="(.*?)".*?IMG/g, function(whole, url) {
      url = url.replace('&amp;', '&');
      var balise = '<img src="'+url+'">';
      return balise;
    })
    .split('{bdc}').join('<img src="data:image/png;base64,'+chart64+'" />')
  return html;
}

function getBDC(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = Charts.newDataTable()
  .addColumn(Charts.ColumnType.STRING, "Sprint")
  .addColumn(Charts.ColumnType.NUMBER, "Standard")
  .addColumn(Charts.ColumnType.NUMBER, "Done");

  var sprintDays = ss.getRangeByName('sprintDays').getValues();
  var standard = ss.getRangeByName('standard').getValues();
  var done = ss.getRangeByName('done').getValues();

  for(var i=0; i<sprintDays.length;i++){
    d = done[i][0]
    if(d == ''){
      var d = null;
    }
    data.addRow([sprintDays[i][0], standard[i][0], d]);
  }
  data.build();

  var chart = Charts.newLineChart()
     .setDataTable(data)
     .setLegendPosition(Charts.Position.NONE)
     .setDimensions(800, 400)
     .build();
  return chart;
}

function getBodyExample(){
  return HtmlService.createHtmlOutputFromFile('exampleTemplate').getContent();
}

function getSubjectExample(){
  return '[BNP Market Tools] Daily report {sprintNumber} - Jour {sprintDay} ({date})'
}

function include(filename) {
      return HtmlService.createHtmlOutputFromFile(filename)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .getContent();
    }

function makeid(){
    var text = "";
    var possible = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";

    for( var i=0; i < 10; i++ )
        text += possible.charAt(Math.floor(Math.random() * possible.length));

    return text;
}
