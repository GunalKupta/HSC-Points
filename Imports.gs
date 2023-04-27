let ss = SpreadsheetApp.getActiveSpreadsheet();
let pointssheet = ss.getSheetByName("Masterlist"); // Points Masterlist

var frozenRows = 4;

// NOTE: Form submission triggers version #, not head deployment

// Extract the title and URL from the submission, import the
// sheet into a new tab of the masterlist spreadsheet, then
// hide the new sheet.
function onFormSubmit(e) {
  let vals = e.namedValues;
  if (vals["Spreadsheet URL"] == undefined) {
    return;
  }
  vals = e.values;
  let rawdate = new Date(vals[0].trim());
  rawdate.setFullYear(new Date().getFullYear());
  let date = rawdate.toLocaleDateString();
  let url = vals[4].trim();
  let title = vals[1].trim();
  let social = vals[2].trim();
  let timestamp = new Date(vals[3]);
  let id = String(timestamp.getTime());

  console.log(JSON.stringify(vals));

  let newsheet = importSheet(url, id, title, date);

  let idsSheet = ss.getSheetByName("Event IDs");
  idsSheet.appendRow([date, title, social, id, url]);

  console.log(`Imported sheet ${newsheet.getName()}\nEvent: ${title}`);
  newsheet.hideSheet();

  SpreadsheetApp.flush();

  extendFormula();
  pointssheet.insertColumnAfter(pointssheet.getLastColumn());
}

// Create a new sheet and use built-in IMPORTRANGE function
function importSheet(url, id, title, date) {
  addImportRangePermission(url);
  let newsheet = ss.insertSheet(id);
  let numcols = SpreadsheetApp.openByUrl(url).getLastColumn();
  newsheet
    .getRange("A1")
    .setFormula(
      `=IMPORTRANGE("${url}","A:${String.fromCharCode(64 + numcols)}")`
    );
  newsheet.getRange(1, numcols + 1).setValue(url);
  return newsheet;
}

// Connect donor sheet with masterlist spreadsheet so that IMPORTRANGE
// is authorized
function addImportRangePermission(url) {
  console.log("opening URL: " + url);
  const donorId = SpreadsheetApp.openByUrl(url).getId();
  const ssId = ss.getId();

  const permurl = `https://docs.google.com/spreadsheets/d/${ssId}/externaldata/addimportrangepermissions?donorDocId=${donorId}`;

  console.log("Permission URL: " + permurl);

  const token = ScriptApp.getOAuthToken();

  const params = {
    method: "post",
    headers: {
      Authorization: "Bearer " + token,
    },
    muteHttpExceptions: true,
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

  let formula = `=IF(ISNA(VLOOKUP(R[0]C1,INDIRECT(R4C[0]&"!C:C"),1,false)),0,IF(ISNA(MATCH("pointsEarned",INDIRECT(R4C[0]&"!A1:1"),0)),100,VLOOKUP(R[0]C1,INDIRECT(R4C[0]&"!C:I"),MINUS(MATCH("pointsEarned",INDIRECT(R4C[0]&"!A1:1"),0),MATCH("uin",INDIRECT(R4C[0]&"!A1:1"),0))+1,false)))`;
  console.log("Formula recorded = " + formula);
  let destrange = pointssheet.getRange(
    frozenRows + 1,
    lastcol,
    pointssheet.getLastRow() - frozenRows - 1
  );
  destrange.setFormulaR1C1(formula);
  console.log("Formulas updated");

  // let countformula = `=COUNTIF(R[-4044]C[0]:R[-1]C[0], ">0")`;
  let countformula = pointssheet.getRange("L3932").getFormulaR1C1();
  destrange = pointssheet.getRange(pointssheet.getLastRow(), lastcol);
  destrange.setFormula(countformula);
}

// Clears most of the main points formulas on the master list
// and replaces them with the same formula to force-refresh any changes.
function forceUpdateFormulas() {
  // let formula = `=IF(ISNA(VLOOKUP(R[0]C1,INDIRECT(R4C[0]&"!C:C"),1,false)),0,IF(ISNA(MATCH("pointsEarned",INDIRECT(R4C[0]&"!A1:1"),0)),100,VLOOKUP(R[0]C1,INDIRECT(R4C[0]&"!C:I"),MINUS(MATCH("pointsEarned",INDIRECT(R4C[0]&"!A1:1"),0),MATCH("uin",INDIRECT(R4C[0]&"!A1:1"),0))+1,false)))`;
  let formula = pointssheet.getRange("J5").getFormulaR1C1();
  Logger.log(formula);
  Logger.log(pointssheet.getLastColumn());
  let range = pointssheet.getRange(
    frozenRows + 1,
    10,
    pointssheet.getLastRow() - frozenRows - 1,
    pointssheet.getLastColumn() - 9
  );
  range.setValue("");
  range.setFormulaR1C1(formula);

  let countformula = `=COUNTIF(R[-3926]C[0]:R[-1]C[0], ">0")`;
  let destrange = pointssheet.getRange(
    pointssheet.getLastRow(),
    pointssheet.getLastColumn()
  );
  destrange.setFormula(countformula);
  SpreadsheetApp.flush();
}

function test() {
  Logger.log(pointssheet.getRange("U3931").getFormulaR1C1());
}
