// RemindMe
// A Google App script to remind me of anniversaries. 
// It currently looks 1 week in advance.
//
// Dependencies:
// Access to anniversary calendar feed from BambooHR (see iCal link on home page).
// Spreadsheet with 2 tabs: Names and Config.
//
// Usage:
// The spreadsheet tab Names lists names for which to send notifications.
// The spreadsheet tab Config includes config values like recipients of the notifications.
//
// Triggers:
// Set it up a project trigger to run main once a day.

// Global constants
var SHEET_NAMES = "Names";
var SHEET_CONFIG = "Config";

// Get config values from config sheet
var config = getConfig();

// Main entrypoint
function main() {
  // Get names from sheet
  var names = getNames();
  
  // Get list of notify days
  var notifyDaysList = getNotifyDaysList();
  
  // For each number of notify days, send notifications
  notifyDaysList.forEach(function(notifyDays) {
    notify(notifyDays, names)
  });
}

function notify(notifyDays, names) {
  // Read events from calendar based on notify days to look ahead
  getEvents(notifyDays).forEach(function(event){
    var params = {
      name: extractNameFromEventTitle(event.getTitle()),
      years: extractYearsFromEventTitle(event.getTitle()),
      date: event.getAllDayStartDate()
    }
    Logger.log("Evaluate: " + params["name"]);
    if (nameExistsInList(params["name"], names)) {
      notifyRecipients(params);
    }
  });
}

// Returns a date object for the specified days in the future
function getFutureDate(daysAhead) {
  var futureDate = new Date();
  futureDate.setDate(futureDate.getUTCDate() + daysAhead);
  return futureDate;
}

// Get events from the calendar
function getEvents(notifyDays) {
  var calendar = CalendarApp.getCalendarsByName(config["CalendarName"])[0];
  return calendar.getEvents(getFutureDate(notifyDays), getFutureDate(notifyDays + 1));
}

// Determines if a name is in the list of names
function nameExistsInList(name, names) {
  return names.indexOf(name) > -1;
}

// Get notify days as a list
function getNotifyDaysList() {
  var notifyDays = config["NotifyDays"];
  return notifyDays.split(",").map(function(notifyDays) {
    return parseInt(notifyDays, 10);
  });
}

// Extracts name of person from event title
// e.g. "Max Mustermann" from "Max Mustermann (1 yr)"
function extractNameFromEventTitle(title) {
  var regex = new RegExp(".*(?=\\s\\()", "g");
  var name = regex.exec(title)[0];
  return name.replace('  ', ' ');
}

// Extracts years from event title
// e.g. "1 yr" from "Max Mustermann (1 yr)"
function extractYearsFromEventTitle(title) {
  var regex = new RegExp("\\(.*\\)", "g");
  var years = regex.exec(title)[0];
  return years.replace('(', '').replace(')', '');
}

// Gets the list of names of interest from spreadsheet
function getNames() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(SHEET_NAMES);
  var numRows = sheet.getLastRow();
  if (numRows == 0) return;
  var names = sheet.getRange(1, 1, numRows, 1).getValues();
  return names.map(function(name) {
    return name[0];
  });
}

// Gets config values from spreadsheet
function getConfig() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(SHEET_CONFIG);
  var numRows = sheet.getLastRow();
  if (numRows == 0) return;
  var values = sheet.getRange(1, 1, numRows, 2).getValues();
  var cfg = [];
  values.forEach(function(value) {
    cfg[value[0]] = value[1];
  });
  return cfg;
}

// Returns a formatted data
// e.g. February 17, 2016
function formatDate(date) {
  return Utilities.formatDate(date, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "MMMM dd, yyyy");
}

// Return the subject of the email
function buildEmailSubject(params) {
  return "Reminder: Anniversary for " + params["name"];
}

// Returns the body of the email
function buildEmailBody(params) {
  return "This is a reminder that " + params["name"] + 
    " is celebrating their " + params["years"] + 
    " anniversary on " + formatDate(params["date"]) + ".";
}

// Sends a notification for the anniversary name
function notifyRecipients(params) {
  var message = {};
  if (config["Debug"]) {
    message = {
      name: config["SenderName"],
      to: config["Debug.RecipientsTo"],
      subject: "[DEBUG] " + buildEmailSubject(params),
      body: buildEmailBody(params)
    };
  }
  else {
    message = {
      name: config["SenderName"],
      to: config["RecipientsTo"],
      cc: config["RecipientsCc"],
      subject: buildEmailSubject(params),
      body: buildEmailBody(params)
    };
  }
  MailApp.sendEmail(message);
}
