# Data-code
Google Apps Script: Spreadsheet Data Updater
Overview
This script is designed to automate the process of updating a Google Sheets document based on CSV data stored in a specific Google Drive folder. The script pulls the necessary counts and updates the sheet with information related to different statuses and assignments, such as "HIF Started", "Cumulative Total", "LLC Applications Submitted", and others. The script handles errors gracefully and provides detailed logs for troubleshooting.

Features
Automatic Updates: Automatically updates the Google Sheet with the latest data from CSV files in a specified folder on Google Drive.

Date Range Handling: Calculates the previous week's date range and ensures headers are created if not present.

Data Counting: Counts the occurrences of specific statuses in CSV files (e.g., "HIF Started", "HIF Signed", etc.) and updates the corresponding cells in the sheet.

Yearly and Cumulative Tracking: Tracks yearly updates and calculates the difference between cumulative values for the current and previous year.

LLC Tracking: Specifically tracks "LLC Applications Submitted" and "LLC Interest Noted" counts from CSV files.

Setup
Required Google Services
Google Sheets API: Access and update Google Sheets.

Google Drive API: Read CSV files stored in a Google Drive folder.

Configuration
To use this script, you will need to configure the following constants:

spreadsheetId: The ID of the Google Sheets file you want to update.

sheetName: The name of the sheet within the Google Sheets file to update.

folderId: The ID of the Google Drive folder containing the relevant CSV files.

Functions
updateSpreadsheetData: Main function that runs the script to update the data in the Google Sheet based on CSV files in the specified Drive folder.

getPreviousWeekDateRangeFormatted: Returns a formatted date range string for the previous week (Monday to Sunday).

ensureDateHeader: Ensures that the header for the current week’s date range is in the sheet, adding it if missing.

updateHIFStartedCount: Updates the count of "HIF Started" records.

updateCumulativeCount: Updates the cumulative total of completed/signed HIF records.

updateEligibleAssignments: Updates the count of eligible HIF assignments.

updateYearlyIncreaseOrDecrease: Tracks the increase or decrease between the current year and the previous year.

updateNotAssignmentEligible: Updates the count of "Not Assignment Eligible" HIF records.

updateLLCApplicationsSubmitted: Updates the count of LLC applications that have been submitted.

updateLLCInterestNoted: Updates the count of LLC interest that has been noted.

Example Usage
Once you’ve configured the constants (spreadsheetId, sheetName, folderId), you can run the updateSpreadsheetData() function. This function will:

Check for the current year and previous week’s date range.

Look for CSV files in the specified folder.

Update the relevant cells in the Google Sheet for various metrics (e.g., HIF Started, Cumulative Counts, LLC Applications Submitted).

Error Handling
The script logs all actions and errors using Logger.log(), which can be reviewed to track down issues and monitor successful updates.

Common Error Messages:
"No CSV files found in the specified folder."

"Could not retrieve 'HIF Started' count from CSV."

"Row for 'HIF Started' not found below year row."

Limitations
The script assumes that CSV files in the folder follow a specific structure with columns like "Status", "EF Paid", and "A&R Registration Date". If the structure changes, adjustments to the script will be necessary.

Only one CSV file will be processed at a time, and the script fetches the first file found in the folder.
