function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Daily scrum')
      .addItem('New Project', 'openProjectForm')
      .addItem('New Sprint', 'openNewSprintForm')
      .addItem('Daily Mail', 'openDailyMailForm')
      .addToUi();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('team');
  if (sheet == null) {
    createTeamSheet();
  }
}

function onInstall(){
  onOpen();
}

function openProjectForm() {
  var html = HtmlService.createHtmlOutputFromFile('newProjectForm')
  .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'New Project');
}

function newProject(formObject) {
  var title = 'BDC'
  if(formObject.projectName != ''){
    title += ' - '+ formObject.projectName
  }
  var ss = SpreadsheetApp.create(title);
  createTeamSheet(ss);
  return ss.getUrl();
}

function createTeamSheet(ss){
  var sheet = ss.getSheets()[0].setName('team');
  sheet.getRange(1,1,1,2).merge().setValue('Tech\' Team');
  sheet.getRange(2,1).setValue('Rôle');
  sheet.getRange(2,2).setValue('Prénom');
  sheet.getRange(3,1).setValue('Architect');
  sheet.getRange(4,1).setValue('Dev');
  sheet.getRange(5,1).setValue('Dev');

  // colors
  sheet.getRange(1,1,1,2).setBackground('#2c7fb8');
  sheet.getRange(2,1,1,2).setBackground('#7fcdbb');
  sheet.getRange(3,1,3,1).setBackground('#edf8b1');

  // size
  sheet.getRange(1,1,1,2).setFontSize(18).setFontWeight('bold');
  sheet.getRange(2,1,1,2).setFontSize(14).setFontWeight('bold');
  sheet.getRange(3,1,3,1).setFontWeight('bold');

  // alignment & borders
  sheet.getRange(1,1,5,2).setHorizontalAlignment('center').setBorder(true, true, true, true, true, true);;
}
