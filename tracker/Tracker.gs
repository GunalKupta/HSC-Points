let lastcol = pointssheet.getLastColumn();
let props = PropertiesService.getDocumentProperties();

function doGet() {
  let output = HtmlService.createTemplateFromFile("tracker/Portal").evaluate();
  output.setTitle("HSC Points Tracker");
  output.addMetaTag("viewport", "width=device-width, initial-scale=1");
  output.setFaviconUrl(
    "https://drive.google.com/uc?id=1sYFRJeZZGg5tXPLEpg-1jyD_nnY3ozAd&export=download&format=png"
  );
  return output;
}

// Validates the entered data and returns a list of events corresponding to
// the UIN
function handleRequest(user) {
  console.log("Handling request");

  let validatedUserRow = validateUser(user);
  if (!validatedUserRow) {
    return { valid: false };
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
  console.log("Got UIN row " + uinrow);
  if (!uinrow || uinrow < 5) {
    return null;
  }

  let userRow = pointssheet.getRange(uinrow, 1, 1, lastcol).getValues()[0];
  let source_lastname = userRow[1].toUpperCase();
  if (source_lastname == lastname) {
    return userRow;
  }

  return null;
}

// Return a list of events a user attended based on UIN
function getEventsForUser(userrow) {
  console.log(JSON.stringify(userrow));

  let data = pointssheet.getRange(1, 1, 4, lastcol).getValues();
  let out = { valid: true };

  let uin = userrow[0];
  out.uin = uin;
  out.lastname = userrow[1];
  out.firstname = userrow[2];
  out.fulfilled = userrow[3] == "Yes";

  let dates = data[0];
  let names = data[1];
  let social = data[2];
  let ids = data[3];
  let events = [];

  let rawdate;
  let date;
  let eventname;
  let indirectdata;
  let rowsFromUin;

  // Return a list of event objects

  for (let c = 7; c < userrow.length; c++) {
    // c represents 0-indexed column number
    let points = parseInt(userrow[c]);
    if (!points) {
      continue;
    }

    let id = ids[c];
    let categoryStr = social[c] ? "SOCIAL" : "ACADEMIC";

    if (c <= 8) {
      // Outside Events or Outside Social: find events linked to uin
      let outsidesheet = ss.getSheetByName(id);
      indirectdata = outsidesheet
        .getRange(2, 4, outsidesheet.getLastRow() - 1, 7)
        .getValues();
      rowsFromUin = getRowsFromUin(uin, indirectdata);
      for (let i = 0; i < rowsFromUin.length; i++) {
        if (!rowsFromUin[i][6]) continue;
        rawdate = new Date(rowsFromUin[i][2]);
        rawdate.setFullYear(2023);
        date = rawdate.toLocaleDateString();
        eventname = `${date} - ${rowsFromUin[i][1]} - <b>${categoryStr}</b>`;
        console.log("Outside result: " + eventname);
        events.push({
          date: date,
          event: rowsFromUin[i][1].trim(),
          social: social[c],
          points: 1,
        });
      }
    } else {
      // Regular event: get data from header columns
      rawdate = new Date(dates[c]);
      rawdate.setFullYear(2023);
      date = rawdate.toLocaleDateString();
      eventname = `${date} - ${names[c]} - <b>${categoryStr}</b>`;
      // eventname = names[c].trim()
      if (points > 100) {
        eventname = `${eventname} (x${points / 100})`;
      }

      events.push({
        date: date,
        event: names[c].trim(),
        social: social[c],
        points: points / 100,
      });
    }
  }

  // sort events by date
  events.sort((a, b) => {
    return new Date(a.date) - new Date(b.date);
  });
  // let sortedEvents = []
  // events.forEach(e => {sortedEvents.push(e.event)});

  out.events = events;
  console.log(JSON.stringify(out));

  return out;
}

// Finds all rows in 2D data where the first element equals uin, and returns them
function getRowsFromUin(uin, data) {
  let out = [];
  let lastcol = data[0].length - 1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] == uin && data[i][lastcol]) out.push(data[i]);
  }
  console.log("Got rows from UIN: " + JSON.stringify(out));
  return out;
}

// Convert the first column of the masterlist into a JSON map
// to make searching for a UIN in the sheet much faster
function createLookupFile() {
  let data = pointssheet
    .getRange(frozenRows + 1, 1, pointssheet.getLastRow() - frozenRows)
    .getValues();
  let out = {};
  for (let i = 0; i < data.length; i++) {
    out[data[i][0]] = i + frozenRows + 1;
  }
  let folder = DriveApp.getFolderById("1DIiSrBIFg7oxeywkT7PnUsPnbaqa5TGX"); // HSC Artifacts folder
  let newFile = DriveApp.createFile(
    "member_row_spring2023",
    JSON.stringify(out),
    MimeType.PLAIN_TEXT
  );
  props.setProperty("lookupFile", newFile.getId());
  newFile.moveTo(folder);
  console.log(newFile.getId());
}

let lock = LockService.getDocumentLock();
// Get the next row for logging a request
function setUpLog() {
  let logsheet = ss.getSheetByName("Request Log");
  let lastrow = logsheet.getLastRow();
  lock.waitLock(10000);
  logsheet.insertRowAfter(lastrow);
  let range = logsheet.getRange(lastrow + 1, 1, 1, 5);
  return range;
}

// Log requests that are made using the web app to track usage
function logRequest(req) {
  let range = setUpLog();

  range.setValues([
    [
      new Date().toLocaleString(),
      req.uin,
      req.firstname + " " + req.lastname,
      req.fulfilled,
      JSON.stringify(req.events).replace("},{", `},\n{`),
    ],
  ]);
  range.setBackground("white");
  range.setFontColor("black");
  SpreadsheetApp.flush();
  lock.releaseLock();
}

function logFailure(req) {
  let range = setUpLog();

  range.setValues([
    [
      new Date().toLocaleString(),
      JSON.stringify(req),
      "",
      "",
      "Invalid Request",
    ],
  ]);
  range.setBackground("red");
  range.setFontColor("white");
  SpreadsheetApp.flush();
  lock.releaseLock();
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(
    "tracker/" + filename
  ).getContent();
}

function outputLookupFile() {
  console.log(props.getProperty("lookupFile"));
}
