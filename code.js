function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  ui.createMenu('VAPL Formulas')
    .addItem('Run This First Time After\nInstall', 'initial')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('BPMS Formulas')
        .addItem('TAT (add working-hours TAT)', 'TAT') // <-- unified
        .addSeparator()
        .addItem('T-x Formula', 'plannedlead')
        .addSeparator()
        .addItem('Specific Time', 'specificTime')
        .addSeparator()
        .addItem('Show planned only when status is NO', 'tatifno')
        .addSeparator()
        .addItem('Show planned only when status is YES', 'tatifyes')
        .addSeparator()
        .addItem('Set Actual Time', 'createTrigger')
        .addSeparator()
        .addItem('Time Delay Formula', 'timeDelay')
    )
    .addToUi();
}

function initial() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();

  // Opening Time
  var o = ui.prompt('Opening Time', 'Enter as HH:MM (24h), e.g., 10:00', ui.ButtonSet.OK_CANCEL);
  if (o.getSelectedButton() !== ui.Button.OK) return;
  var openStr = (o.getResponseText() || '').trim();
  var openFrac = _parseTimeToFractionOrAlert_(openStr, 'Opening Time');
  if (openFrac == null) return;

  // Closing Time
  var c = ui.prompt('Closing Time', 'Enter as HH:MM (24h), e.g., 18:00', ui.ButtonSet.OK_CANCEL);
  if (c.getSelectedButton() !== ui.Button.OK) return;
  var closeStr = (c.getResponseText() || '').trim();
  var closeFrac = _parseTimeToFractionOrAlert_(closeStr, 'Closing Time');
  if (closeFrac == null) return;

  // Validate open < close (same-day window)
  if (closeFrac <= openFrac) {
    ui.alert('Validation error', 'Closing Time must be after Opening Time on the same day (e.g., 10:00 to 18:00).', ui.ButtonSet.OK);
    return;
  }

  // Working-days pattern (WORKDAY.INTL)
  var w = ui.prompt('Working Days Pattern',
                    'Enter a 7-character pattern of 0/1 for Mon..Sun (1 = nonworking).\nExamples:\n  "0000011" = Sat/Sun off (Mon–Fri work)\n  "0000001" = Sun off (Mon–Sat work)',
                    ui.ButtonSet.OK_CANCEL);
  if (w.getSelectedButton() !== ui.Button.OK) return;
  var wd = (w.getResponseText() || '').trim();
  if (!/^[01]{7}$/.test(wd)) {
    ui.alert('Validation error', 'Working Days pattern must be exactly 7 characters of 0/1 (Mon..Sun).', ui.ButtonSet.OK);
    return;
  }

  // Write setup cells
  var sh = ss.getActiveSheet();
  // A1: now() (kept)
  sh.getRange('A1').setFormula('=NOW()');
  // C1: opening time (as time value)
  sh.getRange('C1').setValue(openFrac).setNumberFormat('HH:mm');
  // D1: closing time (as time value)
  sh.getRange('D1').setValue(closeFrac).setNumberFormat('HH:mm');
  // E1: working-days pattern string
  sh.getRange('E1').setValue("'" + String(wd));

  // Remove the old computed value in B1 (and don’t create any new one)
  sh.getRange('B1').clearContent();

  // Hide row 1 as before
  sh.getRange('1:1').activate();
  sh.hideRows(1);

  // Ensure iterative settings as you had
  iterative();
}

/** Helper: parse "HH:MM" → time fraction; alert on error */
function _parseTimeToFractionOrAlert_(txt, label) {
  var ui = SpreadsheetApp.getUi();
  var m = /^([01]?\d|2[0-3]):([0-5]\d)$/.exec(txt);
  if (!m) {
    ui.alert('Validation error',
      label + ' must be in HH:MM 24-hour format (e.g., 09:30, 17:00). You entered: "' + txt + '"',
      ui.ButtonSet.OK);
    return null;
  }
  var hh = parseInt(m[1], 10), mm = parseInt(m[2], 10);
  return (hh * 60 + mm) / 1440; // Sheets time fraction
}

function TAT() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();

  // Prompt for inputs
  var d = ui.prompt('Timestamp Cell', ' Which timestamp cell do you want to use to calculate the Planned Time. Example: A7', ui.ButtonSet.OK_CANCEL);
  if (d.getSelectedButton() !== ui.Button.OK) return;
  var startRef = (d.getResponseText() || '').trim();

  var t = ui.prompt('TAT (hours) Cell', 'Example: F$5', ui.ButtonSet.OK_CANCEL);
  if (t.getSelectedButton() !== ui.Button.OK) return;
  var tatHoursRef = (t.getResponseText() || '').trim();

  var holRange = 'Holidays!$A:$A300';

  // Build formula with guards and the fixed exact-multiple path
  var formula =
    '=IF(OR(ISBLANK(' + startRef + '),ISBLANK(' + tatHoursRef + ')),"",' +
      'LET(' +
        '_start,' + startRef + ',' +
        '_tat,' + tatHoursRef + '/24,' +
        '_open,$C$1,' +
        '_close,$D$1,' +
        '_wd,$E$1,' +
        '_hol,' + holRange + ',' +

        '_sd,INT(_start),' +
        '_st,MOD(_start,1),' +
        '_iswd,WORKDAY.INTL(_sd-1,1,_wd,_hol)=_sd,' +

        '_first,' +
          'IF(_iswd,' +
             'IF(_st<_open,_sd+_open,' +
                'IF(_st>=_close,WORKDAY.INTL(_sd,1,_wd,_hol)+_open,_start)),' +
             'WORKDAY.INTL(_sd,1,_wd,_hol)+_open),' +

        '_ft,MOD(_first,1),' +
        '_avail,_close - MAX(_ft,_open),' +

        'IF(_tat<=_avail,' +
           '_first + _tat,' +
           'LET(' +
             '_rem1,_tat - _avail,' +
             '_daylen,_close - _open,' +
             '_k,INT(_rem1/_daylen),' +
             '_rem2,MOD(_rem1,_daylen),' +
             // base for partial-day path stays 1+_k
             '_base,WORKDAY.INTL(INT(_first),1+_k,_wd,_hol),' +
             // exact-multiple fix: use _k (not 1+_k) and finish at close
             'IF(_rem2=0,' +
               'WORKDAY.INTL(INT(_first),_k,_wd,_hol) + _close,' +
               '_base + _open + _rem2' +
             ')' +
           ')' +
        ')' +
      '))';

  // Write to current cell and format
  ss.getCurrentCell().setFormula(formula);
  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');

  // Optional: keep your autofill behavior
  var currentCell = ss.getCurrentCell();
  ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  currentCell = ss.getCurrentCell();
  ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  ss.getCurrentCell().copyTo(ss.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
}


function createTrigger() { 
  removeTrigger();
  removeTrigger();
  removeTrigger();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger('onChange_new').forSpreadsheet(ss).onChange().create();  
}

function removeTrigger() {
  var allTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < allTriggers.length; i++) {
    if (allTriggers[i].getUniqueId() == allTriggers[i].getUniqueId()) {
      ScriptApp.deleteTrigger(allTriggers[i]);
      break;
    }
  }
}

function importRangeFormula() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var donorSheet = ui.prompt('URL of the spreadsheet from where \ndata has to be imported', ui.ButtonSet.OK_CANCEL);
  if (donorSheet.getSelectedButton() == ui.Button.OK) donorSheet = donorSheet.getResponseText();

  var tabName = ui.prompt('Name of the sheet/tab from where \ndata has to be imported', ui.ButtonSet.OK_CANCEL);
  if (tabName.getSelectedButton() == ui.Button.OK) tabName = tabName.getResponseText();

  var rangeName = ui.prompt('Enter the range from where \ndata has to be imported.', ui.ButtonSet.OK_CANCEL);
  if (rangeName.getSelectedButton() == ui.Button.OK) rangeName = rangeName.getResponseText();

  ss.getCurrentCell().setFormula('=IMPORTRANGE("' + donorSheet + '","' + tabName + '!' + rangeName + '")');
}

function plannedlead() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var fromDate = ui.prompt('Date Cell in which you want to add lead time', ui.ButtonSet.OK_CANCEL);
  if (fromDate.getSelectedButton() != ui.Button.OK) return;
  fromDate = fromDate.getResponseText();

  var leadtime = ui.prompt('Lead Time Cell', ui.ButtonSet.OK_CANCEL);
  if (leadtime.getSelectedButton() != ui.Button.OK) return;
  leadtime = leadtime.getResponseText();

  var daysBefore = ui.prompt('Number of Days Before Lead Time', ui.ButtonSet.OK_CANCEL);
  if (daysBefore.getSelectedButton() != ui.Button.OK) return;
  daysBefore = daysBefore.getResponseText();

  ss.getCurrentCell().setFormula('=IF(' + leadtime + ',' + fromDate + '+' + leadtime + '-' + daysBefore + ',"")');
  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
  var currentCell = ss.getCurrentCell();
  ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  currentCell = ss.getCurrentCell();
  ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  ss.getCurrentCell().copyTo(ss.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
}

function specificTime() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var fromDate = ui.prompt('Date Cell', ui.ButtonSet.OK_CANCEL);
  if (fromDate.getSelectedButton() != ui.Button.OK) return;
  fromDate = fromDate.getResponseText();

  var daysAfter = ui.prompt('Number of Days after previous planned (Write 0 if same day)', ui.ButtonSet.OK_CANCEL);
  if (daysAfter.getSelectedButton() != ui.Button.OK) return;
  daysAfter = daysAfter.getResponseText();

  var tod = ui.prompt('Time of day in hour/24 format', ui.ButtonSet.OK_CANCEL);
  if (tod.getSelectedButton() != ui.Button.OK) return;
  tod = tod.getResponseText();

  ss.getCurrentCell().setFormula('=IF(' + fromDate + ',WORKDAY.INTL(INT(' + fromDate + '),' + daysAfter + ',"0000001",Holidays!$A:$A300)+' + tod + ',"")');
  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
  var currentCell = ss.getCurrentCell();
  ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  currentCell = ss.getCurrentCell();
  ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  ss.getCurrentCell().copyTo(ss.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
}

function actualTime() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var leadtime = ui.prompt('Status Cell', ui.ButtonSet.OK_CANCEL);
  if (leadtime.getSelectedButton() != ui.Button.OK) return;
  leadtime = leadtime.getResponseText();

  var currcella1 = ss.getCurrentCell().getA1Notation();
  ss.getCurrentCell().setFormula('=IF(' + currcella1 + ',' + currcella1 + ',IF(' + leadtime + '<>"",$A$1,""))');
  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
  var currentCell = ss.getCurrentCell();
  ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  currentCell = ss.getCurrentCell();
  ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  ss.getCurrentCell().copyTo(ss.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
}

function timeDelay() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var fromDate = ui.prompt('Planned Cell', ui.ButtonSet.OK_CANCEL);
  if (fromDate.getSelectedButton() != ui.Button.OK) return;
  fromDate = fromDate.getResponseText();

  var leadtime = ui.prompt('Actual Cell', ui.ButtonSet.OK_CANCEL);
  if (leadtime.getSelectedButton() != ui.Button.OK) return;
  leadtime = leadtime.getResponseText();

  ss.getCurrentCell().setFormula('=IF(' + fromDate + ',IF(' + leadtime + '<>"",IF(' + leadtime + '>' + fromDate + ',' + leadtime + '-' + fromDate + ',""),$A$1-' + fromDate + '),"")');
  ss.getActiveRangeList().setNumberFormat('[h]:mm:ss');

  // (Retaining your CF stack)
  var spreadsheet = ss;
  var conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setRanges([spreadsheet.getActiveRange()])
    .whenCellNotEmpty()
    .setBackground('#B7E1CD')
    .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
    .setRanges([spreadsheet.getActiveRange()])
    .whenFormulaSatisfied('=IF(' + leadtime + ',IF(' + leadtime + '>' + fromDate + ',1,0),0)')
    .setBackground('#B7E1CD')
    .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
    .setRanges([spreadsheet.getActiveRange()])
    .whenFormulaSatisfied('=IF(' + leadtime + ',IF(' + leadtime + '>' + fromDate + ',1,0),0)')
    .setBackground('#B7E1CD')
    .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
    .setRanges([spreadsheet.getActiveRange()])
    .whenFormulaSatisfied('=IF(' + leadtime + ',IF(' + leadtime + '>' + fromDate + ',1,0),0)')
    .setBackground('#B7E1CD')
    .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
    .setRanges([spreadsheet.getActiveRange()])
    .whenFormulaSatisfied('=IF(' + leadtime + ',IF(' + leadtime + '>' + fromDate + ',1,0),0)')
    .setBackground('#B7E1CD')
    .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
    .setRanges([spreadsheet.getActiveRange()])
    .whenFormulaSatisfied('=IF(' + leadtime + ',IF(' + leadtime + '>' + fromDate + ',1,0),0)')
    .setBackground('#F4C7C3')
    .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.push(SpreadsheetApp.newConditionalFormatRule()
    .setRanges([spreadsheet.getActiveRange()])
    .whenCellNotEmpty()
    .setBackground('#B7E1CD')
    .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
    .setRanges([spreadsheet.getActiveRange()])
    .whenFormulaSatisfied('=IF(' + leadtime + ',0,IF(' + fromDate + '<$A$1,1,0))')
    .setBackground('#B7E1CD')
    .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule(
  ).setRanges([spreadsheet.getActiveRange()])
    .whenFormulaSatisfied('=IF(' + leadtime + ',0,IF(' + fromDate + '<$A$1,1,0))')
    .setBackground('#B7E1CD')
    .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
    .setRanges([spreadsheet.getActiveRange()])
    .whenFormulaSatisfied('=IF(' + leadtime + ',0,IF(' + fromDate + '<$A$1,1,0))')
    .setBackground('#B7E1CD')
    .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, SpreadsheetApp.newConditionalFormatRule()
    .setRanges([spreadsheet.getActiveRange()])
    .whenFormulaSatisfied('=IF(' + leadtime + ',0,IF(' + fromDate + '<$A$1,1,0))')
    .setBackground('#FCE8B2')
    .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);

  var currentCell = ss.getCurrentCell();
  ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  currentCell = ss.getCurrentCell();
  ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  ss.getCurrentCell().copyTo(ss.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
}

function tatifno() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var fromDate = ui.prompt('Status Cell', ui.ButtonSet.OK_CANCEL);
  if (fromDate.getSelectedButton() != ui.Button.OK) return;
  fromDate = fromDate.getResponseText();
  var formulaincell = ss.getCurrentCell().getFormula().substr(1);
  ss.getCurrentCell().setFormula('=IF(' + fromDate + '="No",' + formulaincell + ',"")');
  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
  var currentCell = ss.getCurrentCell();
  ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  currentCell = ss.getCurrentCell();
  ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  ss.getCurrentCell().copyTo(ss.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
}

function tatifyes() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var fromDate = ui.prompt('Status Cell', ui.ButtonSet.OK_CANCEL);
  if (fromDate.getSelectedButton() != ui.Button.OK) return;
  fromDate = fromDate.getResponseText();
  var formulaincell = ss.getCurrentCell().getFormula().substr(1);
  ss.getCurrentCell().setFormula('=IF(' + fromDate + '="Yes",' + formulaincell + ',"")');
  ss.getActiveRangeList().setNumberFormat('dd/MM/yyyy HH:mm:ss');
  var currentCell = ss.getCurrentCell();
  ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  currentCell = ss.getCurrentCell();
  ss.getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
  currentCell.activateAsCurrentCell();
  ss.getCurrentCell().copyTo(ss.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
}

function iterative() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setRecalculationInterval(SpreadsheetApp.RecalculationInterval.ON_CHANGE);
  spreadsheet.setIterativeCalculationEnabled(true);
  spreadsheet.setMaxIterativeCalculationCycles(1);
  spreadsheet.setIterativeCalculationConvergenceThreshold(0.05);
}

function onChange_new() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var sheet=ss.getActiveSheet();
  var active=ss.getActiveCell();
  var row=active.getRow();
  var col=active.getColumn();
  var val=active.getValue();
  if(val=="Done" ){var ts=sheet.getRange(row, col-1).getValue(); if(ts==""){sheet.getRange(row, col-1).setValue(new Date())}}
  if(val=="Yes"){var ts2=sheet.getRange(row, col-1).getValue(); if(ts2==""){sheet.getRange(row, col-1).setValue(new Date())}}
  if(val=="No" ){var ts3=sheet.getRange(row, col-1).getValue(); if(ts3==""){sheet.getRange(row, col-1).setValue(new Date())}}
  if(val===true ){var ts4=sheet.getRange(row, col-1).getValue(); if(ts4==""){sheet.getRange(row, col-1).setValue(new Date())}}
}