var inSheetId = 1890077395;
var outSheetId = 915915509;

function reset() {
  // Delete triggers
  var trigger = getProjectTriggersByName('checkNewRows');
  if (trigger.length)
    ScriptApp.deleteTrigger(trigger[0]);
  
  // Reset properties
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty("resultsGuardianContactDetails_numberOfRows", 1);
}

function setup() {
  // Trigger every 5 minutes
  ScriptApp.newTrigger('checkNewRows')
      .timeBased()
      .everyMinutes(5)
      .create();
}

function checkNewRows() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var oldRows = parseInt(scriptProperties.getProperty("resultsGuardianContactDetails_numberOfRows") || 1); 
  var newRows = getSheetById(inSheetId).getLastRow();
  if (newRows > oldRows) {
    scriptProperties.setProperty("resultsGuardianContactDetails_numberOfRows", newRows);
    for (var i = oldRows + 1; i <= newRows; i++) {
      processNewRow(i);
    }
  } 
}

function processNewRow(inRow) {
  Logger.log('processNewRow(%s)', inRow);
  var inValues = getSheetById(inSheetId).getSheetValues(inRow, 1, 1, 8)[0];
  var outSheet = getSheetById(outSheetId);
  outSheet.appendRow([""].concat(inValues));
  var outRow = outSheet.getLastRow();
  outSheet.getRange(outRow, 9).setNumberFormat('dddd, d mmmm'); //Guardian deadline
  var dataFormulas = []; //formulas for the columns prefixed "DATA"
  dataFormulas.push(makeFormulaQueryString(0, -8)); //DATA Organisation
  dataFormulas.push(makeFormulaQueryString(0, -8)); //DATA Main contact
  dataFormulas.push(makeFormulaQueryString(0, -7)); //DATA Milestone
  dataFormulas.push(makeFormulaQueryString(0, -7)); //DATA Participant
  dataFormulas.push(makeFormulaQueryString(0, -7)); //DATA Guardian
  dataFormulas.push('=CONCATENATE(\'DO-NOT-TOUCH-urls\'!$B$1,\'DO-NOT-TOUCH-urls\'!$B$7,"?organisation=",R[0]C[-5],"&main_contact=",R[0]C[-4],"&email=",R[0]C[-11],"&milestone=",R[0]C[-3],"&participant=",R[0]C[-2],"&guardian=",R[0]C[-1],"&guardian_email=",R[0]C[-7])'); //URL
  outSheet.getRange(outRow, 10, 1, 6).setFormulas([dataFormulas]);
  outSheet.getRange(outRow, 1).setValues([["YES"]]); //Let Zapier know this row is ready
}

function makeFormulaQueryString(r, c) {
  return '=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(R[' + r + ']C[' + c + '],"%","%25"),"+","%2b")," ","+"),"#","%23"),"&","%26"),"\'","%27"),"’","%27"),"/","%2f")';
}

function getSheetById(id) {
  return SpreadsheetApp.getActive().getSheets().filter(
    function(s) {return s.getSheetId() === id;}
  )[0];
}

function getProjectTriggersByName(name) {
  return ScriptApp.getProjectTriggers().filter(
    function(s) {return s.getHandlerFunction() === name;}
  );
}