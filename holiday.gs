/*================================================================================================================*
  Holiday Tracking by Sheetgo
  ================================================================================================================
  Version:      1.0.0
  Project Page: https://github.com/Sheetgo/holidaytracking
  Copyright:    (c) 2018 by Sheetgo

  License:      GNU General Public License, version 3 (GPL-3.0)
                http://www.opensource.org/licenses/gpl-3.0.html
  ----------------------------------------------------------------------------------------------------------------
  Changelog:
  
  1.0.0  Initial release
 *================================================================================================================*/

// Project Settings
Settings = {

    // The id is a unique set of characters that can be found in the spreadsheet url
    // The spreadsheet id that has the template with your data
    spreadsheetId: SpreadsheetApp.getActiveSpreadsheet().getId(),
   
    // The sheet name
    sheetName: "Holiday Tracking", // Ex.: Holiday Tracking
    requestSheetName: "Time off request form", // Ex.: Time off request form
};


/**
 * Access data from a worksheet to create events in Calendar-related dates
 * Trigger:
 *
 */
function setHolidayCalendar() {
    
    // Access to parameters sheet to calendar_id and event_suffix
    var parameter_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Parameters');
    var calendar_id = parameter_sheet.getRange('A2').getValue();
    var event_suffix = parameter_sheet.getRange('B2').getValue();
  
    // Access Calendar's Spreadsheet
    var spreadsheet = SpreadsheetApp.openById(Settings.spreadsheetId).getSheetByName(Settings.sheetName);
    var calendar = CalendarApp.getCalendarById(calendar_id);

    // The last row and column with data
    var lastRow = spreadsheet.getLastRow();
    var lastColumn = spreadsheet.getLastColumn();

    // Variables with no values that will be used later
    var eventStart, eventEnds, user, day, nextDay, userDays;

    // Get all days
    var dates = spreadsheet.getRange(5, 1, lastRow - 4, 1).getValues();

    // All employees name
    var users = spreadsheet.getRange(1, 4, 1, lastColumn - 3).getValues();

    // Events Range
    var initialDate = dates[0][0];
    var finalDate = dates[dates.length - 1][0];
    initialDate.setDate(initialDate.getDate() + 1)
    finalDate.setDate(finalDate.getDate() + 1)

    // Array to store the days of all users
    var allUsersDays = [];

    // Store All days of all Users
    for (var i = 0; i < users[0].length; i++) {
        allUsersDays.push(spreadsheet.getRange(5, i + 4, lastRow - 4, 1).getValues());
    }

    // Calls the function deleteEvents to delete all events in calendar between the
    // dates with eventSuffix in the title
  
    deleteEvents(calendar, initialDate, finalDate, event_suffix);

    // Initialize consecutive days counter
    var counter = 0;

    // Analyzes the entries of all users to create the Calendar events
    for (var j = 0; j < users[0].length; j++) {

        userDays = allUsersDays[j];
        user = users[0][j];

        // Runs every day
        for (var k = 0; k < dates.length; k++) {

            // Current day
            day = userDays[k][0];

            //If k is greater than or equal to the number of days, there is no next day
            if (k >= dates.length - 1) {
                nextDay = userDays[k][0];
            } else {
                nextDay = userDays[k + 1][0]
            }
            // If the current day does have an event and the following day does not have an event
            // Creates a single day event
            if (day && !nextDay && !counter) {

                // Event start date 
                eventStart = dates[k][0];

                // Creates a single event without an end date
                // In this case, Calendar will create a single day event
                calendar.createAllDayEvent(user + " " + event_suffix, eventStart);

            } else if (day && nextDay) {

                // If the current day and the next day have an event
                // The counter is incremented
                counter++;

            } else if (counter > 0) {

                // If the counter is greater than zero
                // Event start date

                // Event end time
                eventEnds = dates[k + 1][0];
              
                eventStart.setDate(eventStart.getDate() + 1)
                eventEnds.setDate(eventEnds.getDate() + 1)

                // Create an event with a start and end date, according to the counter
                calendar.createAllDayEvent(user + " " + event_suffix, eventStart, eventEnds);

                // The counter is zeroed, so that it can be used again if necessary 
                counter = 0;
            }
        }
    }
}

/**
 * Access the Calendar and delete all events between the dates and contains the eventSuffix
 * @param calendar {String}
 * @param initialDate {Date}
 * @param finalDate {Date}
 * @param eventSuffix {String}
 */
function deleteEvents(calendar, initialDate, finalDate, eventSuffix) {
    // Receive all of the events between the dates
    var eventsToDelete = calendar.getEvents(initialDate, finalDate, {
        search: eventSuffix
    });
    var x = eventsToDelete;
    // Run through all the events to delete
    for (var i = 0; i < eventsToDelete.length; i++) {

        // Delete events
        eventsToDelete[i].deleteEvent();
    }
}

function runHolidayResponse(){
  var planilhaFunc = SpreadsheetApp.openById(Settings.spreadsheetId);
  var sheet = planilhaFunc.getSheetByName(Settings.requestSheetName);
  var lastRow = sheet.getLastRow();
  var dataRange = sheet.getDataRange();
  
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (i=1; i<data.length; i++) {
    var row = data[i];
    var emailAddress = row[1]; // First column
    // Verifiy if the status of request is different of "Waiting" and empty
    // And if the status of e-mail is different of "Sent"
    // If all of these conditions are true, the e-mail will be sent with the answer
    if(row[0] != "" && row[5] != "Waiting" && row[5] != "" && row[6] != 'Sent'){
      var email_answer = 'Your vacation request (from '+ formatDate(row[2]) + ' to ' + formatDate(row[3]) + ') was <b>'+ row[5] + '</b>';
      if (row[5] == "Denied" && row[7] !== "") {
        email_answer += ' because ' + row[7];
      }
      MailApp.sendEmail(emailAddress, "Day Off Request", "", {'htmlBody': email_answer});
      var range = sheet.getRange(1 + i, 7);
      range.setValue('Sent');
      
      // Set dates on Holiday Calendar
      setHolidayCalendar();
    }
  }
}

function formatDate(date) {
  var date = new Date(date);
  return (date.getMonth() + 1) + '/' + (date.getDate() + 1) + '/' +  date.getFullYear();
}


