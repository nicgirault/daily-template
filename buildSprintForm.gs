function openNewSprintForm() {
  var template = HtmlService.createTemplateFromFile('newSprintForm');

  template.defaultSprintNumber = getNextSprint();
  template.defaultStartDate = getStartDate();
  template.defaultEndDate = getEndDate();
  template.defaultTechTeam = getTechTeam();

  var html = template.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'New Sprint');
}

function getNextSprint(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sprint = 1;
  var sheet = ss.getSheetByName('Sprint #' + sprint);
  while(sheet != null){
    sprint += 1;
    sheet = ss.getSheetByName('Sprint #' + sprint);
  }
  return sprint;
}

function getStartDate(){
  var today = moment();
  var tomorrow = today.add('days', 1);
  return moment(tomorrow).format("YYYY-MM-DD");
}

function getEndDate(){
  var today = moment();
  var nextWeek = today.add('weeks', 1);
  return moment(nextWeek).format("YYYY-MM-DD");
}

function getTechTeam(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('team');
  if(sheet == null) {
    return '';
  }
  else {
    team = [];
    for(var i=3;i<20;i++){
      role = sheet.getRange(i,1).getValue();
      if(role == 'Dev' || role == 'Architect'){
        name = sheet.getRange(i,2).getValue()
        if(name != null && name != ''){
          team.push(name);
        }
      }
    }
    return team.join();
  }
}
