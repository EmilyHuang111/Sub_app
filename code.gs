
function onFormSubmit(e) {
  // Extract form data from the response
  var responses = e.values;
  var firstName = responses[1];
  var lastName = responses[2];
  var absenceDate = responses[3];        
  var periods = responses[4];            
  

  var dateObj = new Date(absenceDate);
  

  var eventTitle = firstName + " " + lastName + " (no substitute) - " + periods;
  
  var calendarId = "c_9c555f38fe3d1531d7924cce8c0bc0067fe55ca986fc9428bd0d1630817e249e@group.calendar.google.com";
  var calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    Logger.log("Calendar not found for ID: " + calendarId);
    return;
  }
  
  var event = calendar.createAllDayEvent(eventTitle, dateObj, { description: "Submitted via form" });
  event.setColor(CalendarApp.EventColor.RED);
}


function updateSubstituteColors() {
  var calendarId = "c_9c555f38fe3d1531d7924cce8c0bc0067fe55ca986fc9428bd0d1630817e249e@group.calendar.google.com";
  var calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    Logger.log("Calendar not found for ID: " + calendarId);
    return;
  }
  

  var now = new Date();
  var startRange = new Date(now.getFullYear(), now.getMonth(), now.getDate() - 1);
  var endRange = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 7);
  
  var events = calendar.getEvents(startRange, endRange);
  
  events.forEach(function(event) {
    var title = event.getTitle();
    
    var match = /\(([^)]+)\)/.exec(title);
    
    if (match) {
      var substituteText = match[1].toLowerCase().trim();
      
      if (substituteText === "no substitute" || substituteText === "none") {
        event.setColor(CalendarApp.EventColor.RED);
      } else {
        event.setColor(CalendarApp.EventColor.GREEN);
      }
    } else {
      event.setColor(CalendarApp.EventColor.RED);
    }
  });
}


function sendWeeklySubReport() {
  var calendarId = "c_9c555f38fe3d1531d7924cce8c0bc0067fe55ca986fc9428bd0d1630817e249e@group.calendar.google.com";
  var calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    Logger.log("Calendar not found for ID: " + calendarId);
    return;
  }

  var today = new Date();
  var dayOfWeek = today.getDay();
  var currentMonday;
  if (dayOfWeek === 0) {
    currentMonday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 6);
  } else {
    currentMonday = new Date(today.getFullYear(), today.getMonth(), today.getDate() - (dayOfWeek - 1));
  }
  
  var week1Start = currentMonday;
  var week1End = new Date(currentMonday.getFullYear(), currentMonday.getMonth(), currentMonday.getDate() + 6, 23, 59, 59);
  var week2Start = new Date(currentMonday.getFullYear(), currentMonday.getMonth(), currentMonday.getDate() + 7);
  var week2End = new Date(currentMonday.getFullYear(), currentMonday.getMonth(), currentMonday.getDate() + 13, 23, 59, 59);
  var week3Start = new Date(currentMonday.getFullYear(), currentMonday.getMonth(), currentMonday.getDate() + 14);
  var week3End = new Date(currentMonday.getFullYear(), currentMonday.getMonth(), currentMonday.getDate() + 20, 23, 59, 59);
  
  var overallStart = week1Start;
  var overallEnd = week3End;
  

  var events = calendar.getEvents(overallStart, overallEnd);
  Logger.log("Found " + events.length + " events between " + overallStart + " and " + overallEnd);
  
  var filteredEvents = events.filter(function(event) {
    var title = event.getTitle().toLowerCase();
    return title.indexOf("(no substitute)") !== -1 || title.indexOf("(substitute:") !== -1;
  });
  Logger.log("Filtered events count: " + filteredEvents.length);
  
  var week1Events = [], week2Events = [], week3Events = [];
  var mondayPlus7 = new Date(currentMonday.getTime() + 7 * 24 * 60 * 60 * 1000);
  var mondayPlus14 = new Date(currentMonday.getTime() + 14 * 24 * 60 * 60 * 1000);
  filteredEvents.forEach(function(event) {
    var eventDate = event.getStartTime();
    if (eventDate < mondayPlus7) {
      week1Events.push(event);
    } else if (eventDate < mondayPlus14) {
      week2Events.push(event);
    } else {
      week3Events.push(event);
    }
  });
  
  function sortEventsByStatusAndDate(eventsArray) {
    var noSub = [], hasSub = [];
    eventsArray.forEach(function(event) {
      var title = event.getTitle().toLowerCase();
      if (title.indexOf("no substitute") !== -1) {
        noSub.push(event);
      } else {
        hasSub.push(event);
      }
    });
    noSub.sort(function(a, b) { return a.getStartTime() - b.getStartTime(); });
    hasSub.sort(function(a, b) { return a.getStartTime() - b.getStartTime(); });
    return noSub.concat(hasSub);
  }
  
  week1Events = sortEventsByStatusAndDate(week1Events);
  week2Events = sortEventsByStatusAndDate(week2Events);
  week3Events = sortEventsByStatusAndDate(week3Events);
  
  function buildWeekTable(weekLabel, eventsArray) {
    var table = "<h3>" + weekLabel + "</h3>";
    table += "<table border='1' cellpadding='5' cellspacing='0'>";
    table += "<tr><th>Date</th><th>Teacher</th><th>Substitute</th><th>Periods</th></tr>";
    eventsArray.forEach(function(event) {
      var title = event.getTitle();
      var teacherPart = "";
      var status = "";
      var periods = "";
      var parts = title.split(" - ");
      if (parts.length > 0) { teacherPart = parts[0]; }
      if (parts.length > 1) { periods = parts[1]; }
      var teacherMatch = teacherPart.match(/^(.+?)\s*\((.+)\)$/);
      var teacherName = teacherPart;
      if (teacherMatch) {
        teacherName = teacherMatch[1].trim();
        status = teacherMatch[2].trim();
      }
      var substituteDisplay = "";
      var statusLower = status.toLowerCase();
      if (statusLower.indexOf("no substitute") !== -1) {
        substituteDisplay = "None";
      } else if (statusLower.indexOf("substitute:") !== -1) {
        substituteDisplay = status.replace(/substitute:\s*/i, "");
      } else {
        substituteDisplay = status;
      }
      var eventDate = event.getStartTime().toDateString();
      table += "<tr>";
      table += "<td>" + eventDate + "</td>";
      table += "<td>" + teacherName + "</td>";
      table += "<td>" + substituteDisplay + "</td>";
      table += "<td>" + periods + "</td>";
      table += "</tr>";
    });
    table += "</table>";
    return table;
  }
  
  var message = "<h2>Weekly Substitute Matching Report</h2>";
  message += "<p>Reporting Period: " + overallStart.toDateString() + " to " + overallEnd.toDateString() + "</p>";
  message += buildWeekTable("Current Week (" + week1Start.toDateString() + " to " + week1End.toDateString() + ")", week1Events);
  message += buildWeekTable("Next Week (" + week2Start.toDateString() + " to " + week2End.toDateString() + ")", week2Events);
  message += buildWeekTable("Following Week (" + week3Start.toDateString() + " to " + week3End.toDateString() + ")", week3Events);
  
  // Build the to-do list of events with "no substitute".
  var noSubEvents = filteredEvents.filter(function(event) {
    var title = event.getTitle().toLowerCase();
    return title.indexOf("(no substitute)") !== -1;
  });
  noSubEvents.sort(function(a, b) { return a.getStartTime() - b.getStartTime(); });
  var toDoList = "<h3>To-Do List (No Substitute Assignments)</h3><ul>";
  noSubEvents.forEach(function(event) {
    var title = event.getTitle();
    var teacherPart = "";
    var parts = title.split(" - ");
    if (parts.length > 0) { teacherPart = parts[0]; }
    var teacherMatch = teacherPart.match(/^(.+?)\s*\((.+)\)$/);
    var teacherName = teacherPart;
    if (teacherMatch) { teacherName = teacherMatch[1].trim(); }
    var eventDate = event.getStartTime().toDateString();
    toDoList += "<li>" + eventDate + ": " + teacherName + " (needs substitute assignment)</li>";
  });
  toDoList += "</ul>";
  message += toDoList;
  
  var coordinatorEmail = "emily.huang1113@gmail.com"; 
  var syncSheetUrl = syncCalendarToSheet(coordinatorEmail);  
  
  Logger.log("coordinatorEmail parameter: " + JSON.stringify(coordinatorEmail));
  
  message += "<p>You can view and the live Substitute Sheet here: " +
             "<a href='" + syncSheetUrl + "' target='_blank'>" + syncSheetUrl + "</a></p>";
  
  var subject = "Weekly Substitute Matching Report: " + overallStart.toDateString() + " - " + overallEnd.toDateString();
  MailApp.sendEmail({to: coordinatorEmail, subject: subject, htmlBody: message});
  Logger.log("Weekly report sent to " + coordinatorEmail);
}

function syncCalendarToSheet(coordinatorEmail) {
  var props = PropertiesService.getScriptProperties();
  var sheetId = props.getProperty("CALENDAR_SYNC_SHEET_ID");
  var spreadsheet;

  if (sheetId) {
    try {
      spreadsheet = SpreadsheetApp.openById(sheetId);
      Logger.log("Opened existing spreadsheet with id: " + sheetId);
    } catch (e) {
      Logger.log("Failed to open spreadsheet with id: " + sheetId + ". Creating a new one.");
      sheetId = null;
    }
  }
  if (!sheetId) {
    spreadsheet = SpreadsheetApp.create("Calendar Sync Data");
    sheetId = spreadsheet.getId();
    props.setProperty("CALENDAR_SYNC_SHEET_ID", sheetId);
    Logger.log("Created new spreadsheet with id: " + sheetId);
    if (coordinatorEmail) {
      coordinatorEmail = coordinatorEmail.trim();
      if (coordinatorEmail.indexOf('@') > 0) {
        try {
          spreadsheet.addEditor(coordinatorEmail);
          Logger.log("Added coordinator as editor: " + coordinatorEmail);
        } catch (e) {
          Logger.log("Failed to add editor with email '" + coordinatorEmail + "': " + e);
        }
      } else {
        Logger.log("Provided coordinatorEmail is not a valid email: " + coordinatorEmail);
      }
    } else {
      Logger.log("No coordinatorEmail provided for sharing.");
    }
  }
  
  var sheetName = "CalendarData";
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    Logger.log("Created new sheet named: " + sheetName);
  }
  sheet.clear();

  var headers = ["Name", "Substitute", "Date", "Periods", "Event ID"];
  sheet.appendRow(headers);
  Logger.log("Wrote headers: " + headers.join(", "));
  
  // Retrieve calendar events.
  var calendarId = "c_9c555f38fe3d1531d7924cce8c0bc0067fe55ca986fc9428bd0d1630817e249e@group.calendar.google.com";
  var calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    Logger.log("Calendar not found for ID: " + calendarId);
    return "Error: Calendar not found";
  }
  
 var now = new Date();
var startDate = new Date(now.getFullYear(), now.getMonth(), now.getDate());
var endDate = new Date(9999, 11, 31); 

var events = calendar.getEvents(startDate, endDate);
  Logger.log("Total events fetched: " + events.length);
  
  events.forEach(function(event) {
    var title = event.getTitle();
    var parts = title.split(" - ");
    var teacherPart = parts.length > 0 ? parts[0].trim() : "";
    var periods = parts.length > 1 ? parts[1].trim() : "";
    
    var teacherMatch = teacherPart.match(/^(.+?)\s*\((.+)\)$/);
    var name = teacherPart;
    var substitute = "";
    if (teacherMatch) {
      name = teacherMatch[1].trim();
      substitute = teacherMatch[2].trim();
      if (substitute.toLowerCase().indexOf("no substitute") !== -1) {
        substitute = "None";
      } else if (substitute.toLowerCase().indexOf("substitute:") !== -1) {
        substitute = substitute.replace(/substitute:\s*/i, "").trim();
      }
    }
    
    var eventDate = event.getStartTime().toDateString();

    var eventId = event.getId();
    
    sheet.appendRow([name, substitute, eventDate, periods, eventId]);
  });
  
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow > 1) {
    var dataRange = sheet.getRange(2, 1, lastRow - 1, lastColumn);
    var dataValues = dataRange.getValues();
  
    dataValues.forEach(function(row, index) {
      var subValue = row[1];
      var bgColor = ("" + subValue).trim().toLowerCase() === "none" ? "#ffcccc" : "#ccffcc";
      sheet.getRange(index + 2, 1, 1, lastColumn).setBackground(bgColor);
    });
  }
  
  sheet.hideColumn(sheet.getRange("E1"));
  
  Logger.log("Sync Spreadsheet URL: " + spreadsheet.getUrl());
  return spreadsheet.getUrl();
}

