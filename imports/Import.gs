// COLUMNS in Points Masterlist
// A: UIN
// B: Last Name
// C: First Name
// D: Fulfilled Requirements - Yes/NO
// E: Total points
// F: Total ACADEMIC points
// G: Total SOCIAL points
// H: Total Submission Form points
// I: Extra Social points
// J onward: Individual event reports

let ss = SpreadsheetApp.getActiveSpreadsheet();
let pointssheet = ss.getSheetByName("Points");  // Points Masterlist

// NOTE: Form submission triggers version #, not head deployment

// Extract the title and URL from the submission, import the
// sheet into a new tab of the masterlist spreadsheet, then
// hide the new sheet.
function onFormSubmit(e) {
  let vals = e.values
  let title = vals[3]
  let url = vals[2]

  let newsheet = importSheet(url, title);

  console.log("Imported sheet " + newsheet.getName());
  newsheet.hideSheet();
  
  SpreadsheetApp.flush();

  extendFormula();

  pointssheet.insertColumnAfter(pointssheet.getLastColumn());
}

// Create a new sheet and use built-in IMPORTRANGE function
function importSheet(url, title) {
  let newsheet = ss.insertSheet(title);
  newsheet.getRange("A1").setFormula(`=IMPORTRANGE("${url}", "A:H")`);
  addImportRangePermission(url);
  return newsheet;
}

// Connect donor sheet with masterlist spreadsheet so that IMPORTRANGE
// is authorized
function addImportRangePermission(url) {
  const donorId = SpreadsheetApp.openByUrl(url).getId();
  const ssId = ss.getId();

  const permurl = `https://docs.google.com/spreadsheets/d/${ssId}/externaldata/addimportrangepermissions?donorDocId=${donorId}`;

  console.log("Permission URL: " + permurl);

  const token = ScriptApp.getOAuthToken();

  const params = {
    method: 'post',
    headers: {
      Authorization: 'Bearer ' + token,
    },
    muteHttpExceptions: true
  };
  
  let resp = UrlFetchApp.fetch(permurl, params);
  console.log("Added permissions: " + JSON.stringify(resp));
}

// Fill in formulas in the last column (corresponding to imported event)
// The formula uses VLOOKUP to find the relevant UIN in an imported
// sheet, and return the number of points the member has earned.
function extendFormula() {
  let lastcol = pointssheet.getLastColumn();
  console.log("Last col = " + lastcol);
  let formula = `=IF(ISNA(VLOOKUP(R[0]C1,INDIRECT(R2C[0]&"!C2:H"),6,false)),0,VLOOKUP(R[0]C1,INDIRECT(R2C[0]&"!C2:H"),6,false))`;
  console.log("Formula recorded = " + formula);
  let destrange = pointssheet.getRange(4, lastcol, pointssheet.getLastRow()-3);
  destrange.setFormulaR1C1(formula);
  console.log("Formulas updated");
}

// Clears most of the main points formulas on the master list
// and replaces them with the same formula to force-refresh any changes.
function forceUpdateFormulas() {
  // let formula = `=IF(ISNA(VLOOKUP(R[0]C1,INDIRECT(R2C[0]&"!C2:H"),6,false)),0,VLOOKUP(R[0]C1,INDIRECT(R2C[0]&"!C2:H"),6,false))`;
  let formula = pointssheet.getRange("J4").getFormulaR1C1();
  Logger.log(formula);
  let range = pointssheet.getRange(4, 10, pointssheet.getLastRow()-3, pointssheet.getLastColumn()-9);
  range.setValue("");
  range.setFormulaR1C1(formula);
  SpreadsheetApp.flush();
}

// Deletes all the imported sheets and reimports them
function resetImports() {
  let data = ss.getSheetByName("Imports").getDataRange().getValues();
  for (let row = 1; row < data.length; row++) {
    let url = data[row][2];
    let title = data[row][3];
    let sheet = ss.getSheetByName(title);
    Logger.log(sheet.getName());
    ss.deleteSheet(sheet);

    let newsheet = importSheet(url, title)
    newsheet.hideSheet();
  }

}

