var startRow = 8;
var endRow = 1000;

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Sync to Calendar")
    .addItem("Sync events", "syncEvents")
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("Delete")
        .addItem("Clear removed events", "clearRemovedEvents")
    )
    .addToUi();
}

function syncEvents() {
  var spreadsheet = SpreadsheetApp.getActiveSheet();

  var calendarId = spreadsheet.getRange("H5").getValue();
  var eventCal = CalendarApp.getCalendarById(calendarId);

  var targetRange = spreadsheet.getRange(`A${startRow}:I${endRow}`);
  var data = targetRange.getValues();

  var today = new Date();

  for (var i = 0; i < data.length; i++) {
    var status = data[i][0];
    var project = data[i][1];
    var campaign = data[i][2];
    var description = data[i][3];
    var link = data[i][4];
    var environment = data[i][5];
    var startTime = data[i][6];
    var endTime = data[i][7];
    var eventId = data[i][8];

    if (
      !status ||
      status === "" ||
      !project ||
      project === "" ||
      !campaign ||
      campaign === ""
    ) {
      break;
    }

    if (status !== "On-going") {
      continue;
    }

    try {
      var eventTitle = project + ": " + campaign;
      var eventDescription =
        "project: " +
        project +
        "\ncampaign: " +
        campaign +
        "\ndescription" +
        description +
        "\nlink: " +
        link +
        "\nenvironment: " +
        environment;

      var startTimeDate = new Date(startTime);
      if (!startTime || startTime === "") {
        startTimeDate = today;
      }

      var endTimeDate = new Date(endTime);
      if (!endTime | (endTime === "")) {
        endTimeDate = getLastDateOfTheNextYear(startTimeDate);
      }

      // update finish
      if (today.getDate() > endTimeDate.getDate()) {
        data[i][0] = "Finished";
      }

      // update date
      data[i][6] = startTimeDate;
      data[i][7] = endTimeDate;

      // edit event
      if (eventId) {
        var event = eventCal.getEventById(eventId);
        event.setTitle(eventTitle);
        event.setDescription(eventDescription);
        event.setTime(startTimeDate, endTimeDate);
        event.setTag("project", project);
        event.setTag("environment", environment);
        continue;
      }

      // new event
      var event = eventCal.createEvent(eventTitle, startTimeDate, endTimeDate, {
        description: eventDescription,
        sendInvites: true,
      });
      event.setTag("project", project);
      event.setTag("environment", environment);
      data[i][8] = event.getId();
    } catch (e) {
      alert(`error on project ${project} campaign ${campaign}: ${e.message}`);
    }
  }

  targetRange.setValues(data);
  toast("Updated Calendar");
}

function clearRemovedEvents() {
  var spreadsheet = SpreadsheetApp.getActiveSheet();

  var calendarId = spreadsheet.getRange("H5").getValue();
  var eventCal = CalendarApp.getCalendarById(calendarId);

  var targetRange = spreadsheet.getRange(`A${startRow}:I${endRow}`);
  var data = targetRange.getValues();

  var eventIds = [];

  for (var i = 0; i < data.length; i++) {
    var status = data[i][0];
    var project = data[i][1];
    var campaign = data[i][2];
    var eventId = data[i][8];

    if (
      !status ||
      status === "" ||
      !project ||
      project === "" ||
      !campaign ||
      campaign === ""
    ) {
      break;
    }

    if (eventId !== "") {
      eventIds.push(eventId);
    }
  }

  var events = eventCal.getEvents(
    new Date("01/01/2023"),
    new Date("01/01/2025")
  );
  for (var i = 0; i < events.length; i++) {
    var event = events[i];

    if (!eventIds.includes(event.getId())) {
      event.deleteEvent();
    }
  }
}

function alert(message) {
  SpreadsheetApp.getUi().alert(message);
}

function toast(message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message);
}

function getLastDateOfTheNextYear(date) {
  return new Date(date.getFullYear() + 1, 11, 31);
}
