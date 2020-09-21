const DYNAMIC_VALIDATION_MARKER = "@MajListeDynamique"; // must be written into a note in first line of column
const NB_ROW_TO_HANDLE = 500; // must be a maximum to avoid script failure by timeOut
/*
USAGE
=====
the dynamic validation list is a validation list 'from range'
each validation list refers to a different range, which is calculated from the value from which the list depends
This script help to set the ranges for each line
Each column that contains dynamic dropdown list must be marked with a note containg the marker in the topmost cell
The topmost validation list must refers to the correct range
Then the script will ofsset this range for each validation list below this topmost one (like a formula copy)

If you want to add more lines, copy the last line with validation list as many time as needed below the existing ones
and run the script by giving the line number with which to start (it will take the line just before as template)

*/


/**
 * Look for columns with DynamicValidationMarker in the first cell
 * In this column, looks for a data validation rule based on a range
 * Then all cell below this one containing validation rule are updated by offsetting
 * the data range by one line lower 
 */
function updateDynamicValidationList() {
  UpdateDynamicValidationForColumns(getColumnWithValidationListToUpdate());
}


/**
 * returns an array of column index (1='A') that contains dynamic validation list
 */
function getColumnWithValidationListToUpdate(){
  let table = activeSheet().getDataRange();
  let l = table.getRow();
  var colToUpdate = [];
  for(c=1; c <= table.getWidth(); c++){
    let cell = table.getCell(l, c);
    if(cell.getNote().includes(DYNAMIC_VALIDATION_MARKER)){
      colToUpdate.push(cell.getColumn());
    }
  }
  return colToUpdate;
}


/**
 * 
 * @param {*} col : index of column needing update (1='A')
 */
function UpdateDynamicValidationForColumns(colToUpdate) {
  if (!colToUpdate.length) { return }
  let letter = /[A-Z]+/;
  let sh = activeSheet();
  let startRow = askUserStartRow();
  if (isNaN(startRow)) { return; }
  let maxRow = Math.min(startRow + NB_ROW_TO_HANDLE, sh.getLastRow());
  let range = sh.getRange(startRow - 1, colToUpdate[0],
    maxRow - startRow + 1,
    colToUpdate[colToUpdate.length - 1] - colToUpdate[0] + 1);
  var rules = range.getDataValidations();
  var dataRange, rule, criteria, args, target;
  for (var col = 0; col < rules[0].length; col++) {
    if (colToUpdate.indexOf(col + colToUpdate[0]) < 0) { continue }
    toast("Updating column " + (col + colToUpdate[0]));
    for (var j = 0; j < rules.length; j++) {
      rule = rules[j][col]
      criteria = rule && rule.getCriteriaType() || null;
      if (criteria == 'VALUE_IN_RANGE') { break }
    } // reference validation has been found, or max line reached
    args = rule && rule.getCriteriaValues() || null;
    dataRange = args && args[0] || null;

    for (var line = j + 1; line < rules.length; line++) {
      rule = rules[line][col];
      criteria = rule && rule.getCriteriaType() || null;
      try {
        args = rule && rule.getCriteriaValues() || null;
      } catch (error) {
        alert("exception col " + (col + colToUpdate[0]) + "ligne " + (line + 1) + " Probablement une reference erronée dans un format de validation"
          + " Rien n'est modifié, le script abondonne");
        return;
      }
      if (!rule || (criteria != 'VALUE_IN_RANGE')) { break } // no more validation in this column
      dataRange = dataRange.offset(1, 0);
      if ((line % 100) == 0) {
        toast("range for line " + line + " / " + rules.length + "  :  " + (dataRange && dataRange.getA1Notation()));
      }

      if (args && (args[0] != dataRange)) {
        args[0] = dataRange; // change the range of data for the dropdown list
        rules[line][col] = rule.copy().withCriteria(criteria, args).build();
      }
    } //end loop over lines
  } // end loop over columns
  toast("write update...");
  range.setDataValidations(rules);
  toast("completed until line " + maxRow);
  recordLastRowAchieved(maxRow);
}

/**
 * record the value in the doc properties
 * @param {*} maxRow 
 */
function recordLastRowAchieved(maxRow){
  let docProperties = PropertiesService.getDocumentProperties();
  docProperties.setProperty('maxRowProcessed', maxRow);
}

/**
 * As the user with wich line to start
 * in case of cancel, wrong value , returns NaN
 */
function askUserStartRow(){
  let docProperties = PropertiesService.getDocumentProperties();
  previousMaxRow = PropertiesService.getDocumentProperties().getProperty('maxRowProcessed') || ""
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Envois externes', 'A partir de quelle ligne commencer à incrémenter les zones de liste de validation (dernière traitée : ' + previousMaxRow + ')',
   ui.ButtonSet.OK_CANCEL);
   if (response.getSelectedButton() == ui.Button.OK) {
    return parseInt(response.getResponseText(), 10)
   }
   return NaN;
}

