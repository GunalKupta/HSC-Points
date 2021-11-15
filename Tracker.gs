function doGet() {
  var output = HtmlService.createHtmlOutputFromFile("Portal.html").setTitle("HSC Points Tracker");
  output.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return output;
}

// Check if the user put in the correct UIN, and return a list of events they attended
function getEventsForUser(obj) {
  var out = {valid: false};
  
  // First validate UIN & last name
  var uin = obj.uin.trim();
  var lastname = obj.lastname.toUpperCase();

  var regexp = RegExp(/^\d{9}$/);
  if (!regexp.test(uin)) return out;
  // if (uin.length != 9) return out;

  var lastrow = pointssheet.getLastRow();
  var lastcol = pointssheet.getLastColumn();
  var data = pointssheet.getRange(1,1,lastrow,lastcol).getValues();
  var row;
  for (row = 2; row < data.length; row++) {
    if (data[row][0] == uin && data[row][1].toUpperCase() == lastname) {
      out.valid = true;
      break;
    }
  }
  if (!out.valid) return out;

  // out.lastname = obj.lastname;
  out.firstname = data[row][2];
  out.fulfilled = data[row][3];

  // Return a list of event names
  var userrow = data[row];
  var dates = data[0];
  var names = data[1];
  var social = data[2];
  var events = []

  for (let c = 7; c < userrow.length; c++) {
    let points = parseInt(userrow[c]);
    let date = new Date(dates[c]).toLocaleDateString();
    let eventname = names[c] + " - <b>" + (social[c] ? "SOCIAL" : "ACADEMIC") + "</b>";
    if (date != "Invalid Date") {
      eventname = date + " - " + eventname;
    }
    if (points > 100) {
      events.push(`${eventname} (x${points/100})`);
    } else if (points > 0) {
      events.push(eventname);
    }
  }
  out.events = events;
  // console.log(0, JSON.stringify(dates));
  // console.log(1, JSON.stringify(names));
  console.log(row, JSON.stringify(out));

  return out;
}
