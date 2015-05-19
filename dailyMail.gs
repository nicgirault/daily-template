function sendMail(formObject){
  var draft = GmailApp.getMessageById(formObject.draft);
  var template = draft.getBody();
  var subjectTemplate = draft.getSubject();
  var result = getContentFromTemplate(template, formObject.pointsToValidate);
  var subject = buildSubject(subjectTemplate);
  
  MailApp.sendEmail(draft.getTo(), subject, 'Impossible de lire le contenu', {
    cc: draft.getCc(),
    htmlBody: result[0],
    inlineImages: result[1]
 });
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

function getContentFromTemplate(template, toValidatePoints){
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
    .split('{bdc}').join('<img src="cid:bdc" />')
    .split('{donePoints}').join(''+donePoints)
    .split('{toValidatePoints}').join(''+toValidatePoints)
    .split('{toStandardPoints}').join(''+toStandardPoints)
    .split('{earlyOrLate}').join(earlyOrLate)
    .split('{totalPoints}').join(''+totalPoints)
    .split('{doneColorS}').join('<span style="color: '+doneColor+'">')
    .split('{doneColorE}').join('</span>')
    .split('{validationColorS}').join('<span style="color: '+validationColor+'">')
    .split('{validationColorE}').join('</span>');
 
  var inlineImages = {
    bdc: getBDC().getAs('image/png')
  };
  Logger.log('9');
  
  Logger.log(html);
  html = html.replace(/IMG.*?href="(.*?)".*?IMG/g, function(whole, url) {
    Logger.log(whole);
    url = url.replace('&amp;', '&');
    Logger.log(url);
    var response = UrlFetchApp.fetch(url);
    imgBlob = response.getAs('image/png');
    
    var id = makeid();
    inlineImages[id] = imgBlob;
    
    var balise = '<img src="cid:'+id+'">';
    return balise;
  });

  Logger.log('10');
  return [html, inlineImages];
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
