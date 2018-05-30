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
var Settings = {

    // The id is a unique set of characters that can be found in the spreadsheet url
    // The spreadsheet id that has the template with your data
    spreadsheetId: "<your_spreadsheet_id>",
   
    // The sheet name
    sheetName: '<your_sheet_name>',

    // The id of the Calendar where we are going to create the events
    // if you do not know how to get the Calendar id *
    // READ MORE: #LINK DO BLOGPOST#
    calendarId: "<your_calendar_id>",

    // The suffix of the Calendar event. For example: "John Doe OFF"
    eventSuffix: "<your_event_suffix>"

};


/**
 * Access data from a worksheet to create events in Calendar-related dates
 * Trigger:
 *
 */
function setHolidayCalendar() {

 
    
    // Access Calendar's Spreadsheet
    var spreadsheet = SpreadsheetApp.openById(Settings.spreadsheetId).getSheetByName(Settings.sheetName);
    var calendar = CalendarApp.getCalendarById(Settings.calendarId);

    // The last row and column with data
    var lastRow = spreadsheet.getLastRow();
    var lastColumn = spreadsheet.getLastColumn();

    // Variables with no values that will be used later
    var eventStart, eventEnds, user, day, nextDay, userDays;

    // Get all days
    var dates = spreadsheet.getRange(5, 2, lastRow - 4, 1).getValues();

    // All employees name
    var users = spreadsheet.getRange(1, 4, 1, lastColumn - 3).getValues();

    // Events Range
    var initialDate = dates[0][0];
    var finalDate = dates[dates.length - 1][0];
    finalDate.setDate(finalDate.getDate() + 1)
     
    // Array to store the days of all users
    var allUsersDays = [];

    // Store All days of all Users
    for (var i = 0; i < users[0].length; i++) {
        allUsersDays.push(spreadsheet.getRange(5, i + 4, lastRow - 4, 1).getValues());
    }

    // Calls the function deleteEvents to delete all events in calendar between the
    // dates with eventSuffix in the title
    deleteEvents(calendar, initialDate, finalDate, Settings.eventSuffix);

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
                calendar.createAllDayEvent(user + " " + Settings.eventSuffix, eventStart);

            } else if (day && nextDay) {

                // If the current day and the next day have an event
                // The counter is incremented
                counter++;

            } else if (counter > 0) {

                // If the counter is greater than zero
                // Event start date
                eventStart = dates[k - counter][0];

                // Event end time
                eventEnds = dates[k + 1][0];

                // Create an event with a start and end date, according to the counter
                calendar.createAllDayEvent(user + " " + Settings.eventSuffix, eventStart, eventEnds);

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

    // Run through all the events to delete
    for (var i = 0; i < eventsToDelete.length; i++) {

        // Delete events
        eventsToDelete[i].deleteEvent();
    }
}

