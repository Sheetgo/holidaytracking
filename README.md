# Holiday Tracking

## Introduction

Sheetgo team needs to be flexible with the holiday and vacation planning. We set a fixed number of days per year that an employee can take off and let each individual within the organization to choose which days they want to take off.

This Google Apps Script takes inputs from a [spreadsheet](https://docs.google.com/spreadsheets/d/1oPD22qiVj65Ud1_U-O-44xJekYQ6SFtlyiL_slWydTE/copy) and automatically creates calendar events. The script can be configured to run once a day so that any changes to the spreadsheet are automatically reflected on the calendar.

These require a Google account and an explicit permission, but in some cases may be a good fit.

## How to configure
1. Click the link to create a copy of the spreadsheet holiday tracking [template](https://docs.google.com/spreadsheets/d/1oPD22qiVj65Ud1_U-O-44xJekYQ6SFtlyiL_slWydTE/copy).
2. Fill out the spreadsheet with the name of your teammates and the total days allocated for holidays for each one.
3. Copy and paste the ID of the spreadsheet to a clipboard or text editor as you will need it in the future.
4. Create a new calendar (this is will be the calendar that you’ll share with the team).
5. Copy the calendar ID by selecting Settings and Sharing, scroll down to Integrate Calendar and copy the calendar ID and paste it in a clipboard or text editor for you to use later.
6. Click on Tools and then Script Editor to edit the script.
7. Now just substitute your unique Calendar and Spreadsheet IDs. Paste the ID of the spreadsheet that you copied in step 5 in between the quotes, substituting any existing values. Do the same things with the ID of the calendar that you copied in step 6 and paste in place of the calendar ID.
8. To execute the script manually select the function setHolidayCalendar and then click the run (play) button. You can also create automated triggers to execute the script automatically.
9. To add an automatic trigger open the script and click on the clock icon, select the function setHolidayCalendar and choose the time of your preference.








## Version
+ 1.0.0 - Initial release.

## More information

For detailed information please visit the [blog](https://blog.sheetgo.com/google-cloud-solutions/holiday-tracking/).
