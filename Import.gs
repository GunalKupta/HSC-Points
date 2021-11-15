var ss = SpreadsheetApp.getActiveSpreadsheet();
var pointssheet = ss.getSheetByName("Points");  // Points Masterlist

// NOTE: Form submission triggers version #, not head deployment

// Extract the title and URL from the submission, import the
// sheet into a new tab of the masterlist spreadsheet, then
// hide the new sheet.
function onFormSubmit(e) {
  var vals = e.values
  var title = vals[3]
  var url = vals[2]

  // var newsheet = copySheet(url, title);
  var newsheet = importSheet(url, title);

  console.log("Imported sheet " + newsheet.getName());
  newsheet.hideSheet();
  
  SpreadsheetApp.flush();

  extendFormula();

  pointssheet.insertColumnAfter(pointssheet.getLastColumn());

  pointssheet.sort(4, false);
}

// Create a new sheet and use built-in IMPORTRANGE function
function importSheet(url, title) {
  var newsheet = ss.insertSheet(title);
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
  
  var resp = UrlFetchApp.fetch(permurl, params);
  console.log("Added permissions: " + JSON.stringify(resp));
}

// Fill in formulas to recalculate points
function extendFormula() {
  var lastcol = pointssheet.getLastColumn();
  console.log("Last col = " + lastcol);
  var formula = `=IF(ISNA(VLOOKUP(R[0]C1,INDIRECT(R1C[0]&"!C2:H"),6,false)),0,VLOOKUP(R[0]C1,INDIRECT(R1C[0]&"!C2:H"),6,false))`
  console.log("Formula recorded = " + formula);
  var destrange = pointssheet.getRange(3, lastcol, pointssheet.getLastRow()-2);
  destrange.setFormulaR1C1(formula);
  console.log("Formulas updated");
}

