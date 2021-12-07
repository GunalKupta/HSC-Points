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

let pointsubmission = ss.getSheetByName("Point Submission Form");
let extrasocial = ss.getSheetByName("Extra Social Points");
let lastcol = pointssheet.getLastColumn();
let props = PropertiesService.getDocumentProperties();

function doGet() {
  let output = HtmlService.createTemplateFromFile("tracker/Portal").evaluate();
  output.setTitle("HSC Points Tracker");
  output.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  output.setFaviconUrl("https://drive.google.com/uc?id=1IvmnYJP8lV0uRY1WcBDZ-6e_1qd-O0Rw&export=download&format=png");
  return output;
}

// Validates the entered data and returns a list of events corresponding to
// the UIN
function handleRequest(user) {
  console.log("Handling request");

  let validatedUserRow = validateUser(user);
  if (!validatedUserRow) {
    return {valid: false};
  }

  console.log("Validated user");

  return getEventsForUser(validatedUserRow);
}

// Parses a text file containing a JSON map of UIN -> row in spreadsheet
// https://stackoverflow.com/a/29347372
function findRowInFile(uin) {
  let fileId = props.getProperty("lookupFile");
  let file = DriveApp.getFileById(fileId),
    filedata = file.getBlob().getDataAsString(),
    data = JSON.parse(filedata);
  
  return data[uin];
}

// Given an object containing the user input (uin and lastname),
// check that they match and find the row in the masterlist
// corresponding to the UIN
function validateUser(obj) {
  let uin = obj.uin.trim();
  let lastname = obj.lastname.trim().toUpperCase();

  let regexp = RegExp(/^\d{9}$/);
  if (!regexp.test(uin)) return null;

  let uinNum = parseInt(uin);

  let uinrow = findRowInFile(uinNum);
  console.log("Got UIN row "+uinrow);
  if (!uinrow || uinrow < 4) {
    return null;
  }

  let userRow = pointssheet.getRange(uinrow,1,1,lastcol).getValues()[0];
  
  if (userRow[1].toUpperCase() == lastname) {
    return userRow;
  }

  return null;
}

// Return a list of events a user attended based on UIN
function getEventsForUser(userrow) {
  
  let data = pointssheet.getRange(1,1,3,lastcol).getValues();
  let out = {valid: true};

  let uin = userrow[0];
  out.uin = uin;
  out.lastname = userrow[1];
  out.firstname = userrow[2];
  out.fulfilled = userrow[3];

  let dates = data[0];
  let names = data[1];
  let social = data[2];
  let events = []

  let rawdate;
  let date;
  let eventname;
  let indirectdata;
  let rowsFromUin;

  // Return a list of event names

  for (let c = 7; c < userrow.length; c++) {
    let points = parseInt(userrow[c]);
    if (!points) {
      continue;
    }

    if (c == 7) {
      // Point Submission Form: find events linked to uin
      indirectdata = pointsubmission.getRange(2,3,pointsubmission.getLastRow()-1,7).getValues();
      rowsFromUin = getRowsFromUin(uin, indirectdata);
      for (let i = 0; i < rowsFromUin.length; i++) {
        rawdate = new Date(rowsFromUin[i][4]);
        rawdate.setFullYear(2021);
        date = rawdate.toLocaleDateString();
        eventname = date + " - (Point Submission Form) " + rowsFromUin[i][2] + " - <b>ACADEMIC</b>";
        console.log("Point Submission result: " + eventname);
        events.push(eventname.trim());
      }

    } else if (c == 8) {
      // Extra Social Points: find events linked to uin
      indirectdata = extrasocial.getRange(2,3,extrasocial.getLastRow()-1,6).getValues();
      rowsFromUin = getRowsFromUin(uin, indirectdata);
      for (let i = 0; i < rowsFromUin.length; i++) {
        rawdate = new Date(rowsFromUin[i][2]);
        rawdate.setFullYear(2021);
        date = rawdate.toLocaleDateString();
        eventname = date + " - " + rowsFromUin[i][3] + " - <b>SOCIAL</b>"
        events.push(eventname.trim());
      }

    } else {
      // Regular event: get data from header columns
      eventname = names[c] + " - <b>" + (social[c] ? "SOCIAL" : "ACADEMIC") + "</b>";
      rawdate = new Date(dates[c]);
      rawdate.setFullYear(2021);
      date = rawdate.toLocaleDateString();
      // if (date != "Invalid Date") {
        eventname = date + " - " + eventname;
      // }
      events.push(eventname.trim());
    }
  }
  out.events = events;
  console.log(JSON.stringify(out));

  return out;
}

// Finds all rows in 2D data where the first element is uin, and returns them
function getRowsFromUin(uin, data) {
  let out = []
  let lastcol = data[0].length-1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] == uin && data[i][lastcol]) out.push(data[i]);
  }
  console.log("Got row from UIN: " + JSON.stringify(out));
  return out;
}

// Convert the first column of the masterlist into a JSON map
// to make searching for a UIN in the sheet much faster
function createLookupFile() {
  let data = pointssheet.getRange(4,1,5062-3).getValues();
  let out = {};
  for (let i = 0; i < data.length; i++) {
    out[data[i][0]] = i+4;
  }
  let folder = DriveApp.getFolderById("1DIiSrBIFg7oxeywkT7PnUsPnbaqa5TGX"); // HSC Artifacts folder
  let newFile = DriveApp.createFile("member_row", JSON.stringify(out), MimeType.PLAIN_TEXT);
  props.setProperty("lookupFile", newFile.getId());
  newFile.moveTo(folder);
  console.log(newFile.getId());
}

let lock = LockService.getDocumentLock();
// Get the next row for logging a request
function setUpLog() {
  let logsheet = ss.getSheetByName("Request Log")
  let lastrow = logsheet.getLastRow();
  lock.waitLock(10000);
  logsheet.insertRowAfter(lastrow);
  let range = logsheet.getRange(lastrow+1,1,1,5);
  return range;
}

// Log requests that are made using the web app to track usage
function logRequest(req) {
  let range = setUpLog();

  range.setValues([[new Date().toLocaleString(), req.uin, req.firstname+" "+req.lastname, req.fulfilled, JSON.stringify(req.events)]])
  range.setBackground("white");
  range.setFontColor("black");
  SpreadsheetApp.flush();
  lock.releaseLock();
}

function logFailure(req) {
  let range = setUpLog();

  range.setValues([[new Date().toLocaleString(), JSON.stringify(req), "", "", "Invalid Request"]]);
  range.setBackground("red");
  range.setFontColor("white");
  SpreadsheetApp.flush();
  lock.releaseLock();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile("tracker/"+filename)
    .getContent();
}

function outputLookupFile() {
  console.log(props.getProperty("lookupFile"));
}
