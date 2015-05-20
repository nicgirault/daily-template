var COLOR1 = EDIT_COLOR = '#5ab4ac';
var COLOR2 = '#d8b365';
var GREY = '#f5f5f5';
var COLOR1_LIGHT = '#f6e8c3';
var COLOR2_LIGHT = '#c7eae5';

function createSprintSheet(formObject) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.insertSheet('Sprint #' + formObject.sprintNumber, 0);
  buildNewSprintHeader(formObject);
  buildNewSprintWeek(formObject);
}

function buildNewSprintHeader(formObject){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  sheet = ss.getActiveSheet();

  sheet.getRange('B3').setValue('Célérité');
  sheet.getRange('B4').setValue('Total des ressources');
  sheet.getRange('B5').setValue('Points total sprint');
  ss.setNamedRange('totalPoints', sheet.getRange('C5'));

  // top left table style
  sheet.getRange('C3').setBackground(EDIT_COLOR);
  sheet.getRange('B3:C5').setBorder(true, true, true, true, true, true).setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.setColumnWidth(1, 10).setColumnWidth(2, 200).setRowHeight(1, 10).setRowHeight(2, 40).setRowHeight(3, 25).setRowHeight(4, 25).setRowHeight(5, 25);

  sheet.getRange('F2:K2').setBackground(GREY).setBorder(true, true, true, true, true, true).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('F2').setValue('Sprint');
  sheet.getRange('G2').setValue('#'+formObject.sprintNumber);
  ss.setNamedRange('sprintNumber', sheet.getRange('G2'));

  sheet.getRange('H2').setValue('Du');
  sheet.getRange('I2').setValue(formObject.startDate);
  ss.setNamedRange('startDate', sheet.getRange('I2'));

  sheet.getRange('J2').setValue('au');
  sheet.getRange('K2').setValue(formObject.endDate);
  sheet.getRange('F3:K3').setBackground(COLOR2).setFontColor('white').setFontSize(14).setBorder(true, true, true, true, true, true).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('F3').setValue('Goal');
  sheet.getRange('G3:K3').merge();
  ss.setNamedRange('sprintGoal', sheet.getRange('G3'));

}

function buildNewSprintWeek(formObject){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  var sprint = formObject.sprintNumber;
  var devTeam = formObject.devTeam.split(',');

  // size of the main table
  var rowOffset = 7;
  var colOffset = 2;
  var width = 4 + devTeam.length;

  sheet.getRange(rowOffset, colOffset, 1, width).setBackground('white').setFontColor('black').setFontSize(14).setBorder(true, true, true, true, true, true).setFontWeight('bold');
  sheet.getRange('B7').setValue('Jours');
  sheet.getRange('C7:E7').merge().setValue('Avancement');
  sheet.getRange(rowOffset, colOffset + 4, 1, devTeam.length).merge().setValue('Dev\' Team');
  sheet.getRange(rowOffset, colOffset, 1, width).setBackground(GREY);


  sheet.getRange(rowOffset + 1, colOffset, 1, width).setBackground(COLOR2).setFontColor('white').setFontWeight('bold');
  sheet.getRange('C8').setValue('Standard');
  sheet.getRange('D8').setValue('Done');
  sheet.getRange('E8').setValue('Différence');
  for(var idx in devTeam){
    sheet.getRange(rowOffset + 1, colOffset + 4 + parseInt(idx)).setValue(devTeam[idx]);
  }

  moment.locale('fr', {
    weekdays : 'dimanche_lundi_mardi_mercredi_jeudi_vendredi_samedi'.split('_')
  });
  sheet.getRange('B9').setValue('Départ');
  var start = moment(formObject.startDate);
  var end = moment(formObject.endDate);
  dates = moment().range(start, end);
  nDays = 0;
  dates.by('days', function(date) {
    Logger.log(date.format('dddd'));
    if([1,2,3,4,5].indexOf(date.weekday()) > -1){
      sheet.getRange(rowOffset + 3 + nDays, colOffset).setValue(date.format('dddd'));
      nDays += 1;
    }
  });
  ss.setNamedRange('sprintDays', sheet.getRange(rowOffset + 2, colOffset, nDays + 1));

  resources = sheet.getRange(rowOffset + 2, colOffset + 4, nDays + 1, devTeam.length).setBackground(EDIT_COLOR)
  sheet.getRange('C4').setFormula('=SUM('+resources.getA1Notation()+')');
  sheet.getRange('C5').setFormula('=C3*C4');

  nDays = 0;
  dates.by('days', function(date) {
    if(date.weekday() < 5){
      sheet.getRange(rowOffset + 3 + nDays, colOffset + 1).setFormula('=$C$3*SUM('+sheet.getRange(rowOffset + 3, colOffset + 4, nDays+1, devTeam.length).getA1Notation()+')');
      var std = sheet.getRange(rowOffset + 3 + nDays, colOffset + 1).getA1Notation();
      var done = sheet.getRange(rowOffset + 3 + nDays, colOffset + 2).getA1Notation();
      sheet.getRange(rowOffset + 3 + nDays, colOffset + 3).setFormula('=if(isblank('+done+'),,'+done+'-'+std+')');
      nDays += 1;
    }
  });

  sheet.getRange(rowOffset + 2, colOffset + 2, nDays + 1).setBackground(EDIT_COLOR);
  sheet.getRange(rowOffset, colOffset, nDays + 3, width).setHorizontalAlignment('center').setBorder(true, true, true, true, null, null);

  // set value 0 to the start row
  sheet.getRange(rowOffset+2,colOffset+1,1,width-1).setValue(0);

  bdcPreSeries(rowOffset, colOffset + width + 1, nDays, '$C$5', [9, 3]);
}

function bdcPreSeries(rowOffset, colOffset, nRows, totalPoints, stdRange){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  sheet.getRange(rowOffset, colOffset, nRows + 3, 2).setBackground(GREY).setFontSize(8).setBorder(true, true, true, true, true, true).setFontStyle('italic').setHorizontalAlignment('center');
  var stdFormulas = [];
  var doneFormulas = [];
  for(var i=0; i < nRows+1; i++){
    stdFormulas.push(['='+totalPoints+'-'+sheet.getRange(stdRange[0] + i, stdRange[1]).getA1Notation()])

    var currentDone = sheet.getRange(stdRange[0] + i, stdRange[1] + 1).getA1Notation();
    doneFormulas.push(['=if(isblank('+currentDone+'),,'+totalPoints+'-'+currentDone+')'])
  }

  sheet.getRange(rowOffset, colOffset, 1, 2).merge().setValue('BDC');
  sheet.getRange(rowOffset + 1, colOffset).setValue('Standard');
  sheet.getRange(rowOffset + 1, colOffset + 1).setValue('Done');

  std = sheet.getRange(rowOffset + 2, colOffset, nRows + 1).setFormulas(stdFormulas);
  done = sheet.getRange(rowOffset + 2, colOffset + 1, nRows + 1).setFormulas(doneFormulas);

  ss.setNamedRange('standard', sheet.getRange(rowOffset + 2, colOffset, nRows + 1));
  ss.setNamedRange('done', sheet.getRange(rowOffset + 2, colOffset + 1, nRows + 1));

  data = sheet.getRange(rowOffset + 2, colOffset, nRows + 1, 2);
  drawBDC(data, ss.getRangeByName('sprintDays'));
}

function drawBDC(data, Xaxis){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();

  var chart = sheet.newChart()
     .setChartType(Charts.ChartType.LINE)
     .addRange(Xaxis)
     .addRange(data)
     .setPosition(17, 2, 0, 0)
     .asLineChart()
     .setLegendPosition(Charts.Position.NONE)
     .build();
 sheet.insertChart(chart);
}
