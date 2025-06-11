🚀 Google Apps Script Projects Portfolio
Welcome! This repository contains my collection of Google Apps Script automations tailored to streamline workflows in Google Sheets — from navigation helpers to backup sync tools and automated notifications. Each project is carefully designed to improve efficiency, accuracy, and user experience.

1. 🔍 Accounts Navigator Button
Purpose:
Quickly navigate to and highlight a persona’s account details in the Accounts sheet from the Sales or Contacts sheets.

How it works:
Select a name (e.g., “Naomi”) in the Sales sheet (Column C).

Click the Accounts button.

The script searches the “002_Accounts” sheet for all occurrences of the name.

It highlights those rows in yellow and opens the Accounts sheet in a new tab.

Reverse navigation to Contacts is also supported.

Key script:
javascript
Copy
Edit
function checkAccounts(nameToCheck) {
  var accountsSpreadsheetId = '1nnuUewfzqQX2EZIpg4jDJbSUH14Nlj1wYVtftbD3Ivo'; // ID of the Accounts spreadsheet
  var accountsSheet = SpreadsheetApp.openById(accountsSpreadsheetId).getSheetByName('002_Accounts');

  if (!accountsSheet) {
    Logger.log("The '002_Accounts' sheet does not exist.");
    return;
  }

  var accountsData = accountsSheet.getDataRange().getValues();
  var found = false;

  accountsSheet.getRange(1, 1, accountsSheet.getMaxRows(), accountsSheet.getMaxColumns()).setBackground(null);

  for (var i = 0; i < accountsData.length; i++) {
    if (accountsData[i][1] == nameToCheck) {
      found = true;
      var rowIndex = i + 1;
      accountsSheet.getRange(rowIndex, 1, 1, accountsSheet.getLastColumn()).setBackground('yellow');
    }
  }

  if (found) {
    Logger.log("Name found and highlighted in Accounts.");
    accountsSheet.setActiveRange(accountsSheet.getRange(rowIndex, 1, 1, accountsSheet.getLastColumn()));
    var url = 'https://docs.google.com/spreadsheets/d/' + accountsSpreadsheetId + '/edit#gid=' + accountsSheet.getSheetId();
    var html = HtmlService.createHtmlOutput('<script>window.open("' + url + '");</script>');
    SpreadsheetApp.getUi().showModalDialog(html, 'Open Accounts Sheet');
  } else {
    Logger.log('Name not found in Accounts.');
  }
}

function checkAccounts1Wrapper() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var salesSheet = ss.getSheetByName('014_Sales');
  
  if (!salesSheet) {
    Logger.log("The 'Sales' sheet does not exist.");
    return;
  }
  
  var range = salesSheet.getActiveCell();
  
  if (range.getColumn() == 3) {
    var nameToCheck = range.getValue();
    checkAccounts(nameToCheck);
  } else {
    Logger.log("Please click the button next to a name in Column C.");
  }
}
2. 🔄 Automatic Sales Stage Updater
Purpose:
Automatically sets the sales stage to “Offered” for any new entries with an RLM-Number but no stage set.

How it works:
The script runs through the ‘014_Sales’ sheet rows.

Checks if Column B (RLM-Number) is filled and Column G (Stage) is empty.

Sets the Stage to “Offered”.

Script snippet:
javascript
Copy
Edit
function updateStage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('014_Sales');
  
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const stageColumn = 7; // Column G

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] && !data[i][stageColumn - 1]) {
      sheet.getRange(i + 1, stageColumn).setValue('Offered');
    }
  }
}
3. 📧 Automated Email Notification on Deal Closure
Purpose:
When a sales deal is closed, an email notification is automatically sent to the relevant departments, with a timestamp and status update.

How it works:
Scans ‘014_Sales’ for rows marked “Closed” in the Stage column.

Sends an email to all departments with deal details.

Marks the deal as “Email Sent” and adds a timestamp to avoid duplicate emails.

Can be scheduled via time-driven trigger.

Script snippet:
javascript
Copy
Edit
function sendEmailOnDealClosure() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var emailRecipient = "philipp@rushlake-media.com,ivor@rushlake-media.com,muhiiga@rushlake-media.com,operations@rushlake-media.com";
  var subject = "NEW SALE JUST CLOSED";
  var message = "";
  var now = new Date();

  for (var i = 1; i < data.length; i++) {
    var status = data[i][6];
    var emailSent = data[i][9];
    var timestamp = data[i][10];

    if (status === "Closed" && emailSent === "") {
      var rlmNumber = data[i][0];
      var title = data[i][1];
      var account = data[i][2];
      var licensee = data[i][5];
      var status = data[i][6];

      message += "🔹 RLM-Number: " + rlmNumber + "\n";
      message += "🔹 Title: " + title + "\n";
      message += "🔹 Account: " + account + "\n";
      message += "🔹 Licensee: " + licensee + "\n";
      message += "🔹 Stage: " + status + "\n";
      message += "✅ This Sale is closed.\n\n";

      sheet.getRange(i + 1, 10).setValue("Email Sent");

      if (timestamp === "") {
        sheet.getRange(i + 1, 11).setValue(now);
      }
    }
  }

  if (message !== "") {
    MailApp.sendEmail(emailRecipient, subject, message);
  }
}
4. 📅 Sorted Popup Sheet for Imported Data
Purpose:
Sorts data imported from another spreadsheet by date (Column G) and creates a clean sorted popup sheet.

How it works:
Imports data from a source sheet via URL.

Sorts all rows by date in ascending or descending order.

Writes the sorted data into a new or cleared sheet named ‘SortedPopup’.

Activates the popup sheet for easy viewing.

Script snippet:
javascript
Copy
Edit
function sortImportrangePopupSheet(ascending = false) {
  const sourceUrl = 'https://docs.google.com/spreadsheets/d/1PMF58YB2rAX-ZcBfUpSWpz9_fRLJi1TTDZz0lYUZ-sk/edit';
  const sourceRange = '021_Inbound Delivery!A1:P';
  const sourceSheetName = '021_Inbound Delivery';
  const popupSheetName = 'SortedPopup';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceData = SpreadsheetApp.openByUrl(sourceUrl)
    .getSheetByName(sourceSheetName)
    .getRange(sourceRange)
    .getValues();

  const header = sourceData[0];
  const dataToSort = sourceData.slice(1);
  const dateColumnIndex = 6; // Column G

  dataToSort.sort((a, b) => {
    const dateA = new Date(a[dateColumnIndex]);
    const dateB = new Date(b[dateColumnIndex]);
    if (isNaN(dateA)) return 1;
    if (isNaN(dateB)) return -1;
    return ascending ? dateA - dateB : dateB - dateA;
  });

  const sortedData = [header, ...dataToSort];

  let popupSheet = ss.getSheetByName(popupSheetName);
  if (!popupSheet) {
    popupSheet = ss.insertSheet(popupSheetName);
  } else {
    popupSheet.clearContents();
  }

  popupSheet.getRange(1, 1, sortedData.length, sortedData[0].length).setValues(sortedData);

  ss.setActiveSheet(popupSheet);
}
5. 🔄 Backup Restore for Manual Data Sync
Purpose:
Keeps manual data synchronized between a main sheet and a backup sheet even if sorting or deletions occur.

How it works:
Reads titles from main sheet (‘THE INBOUND’) Column A.

Matches manual data columns (B-G) from the backup sheet to main sheet based on title.

Restores manual data in the exact order of the main sheet titles.

Ensures backup always reflects main sheet changes, even without importrange.

Script snippet:
javascript
Copy
Edit
function restoreManualData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName('THE INBOUND');
  var backupSheet = ss.getSheetByName('Backup');

  var mainTitles = mainSheet.getRange('A2:A' + mainSheet.getLastRow()).getValues().flat();
  var backupData = backupSheet.getRange('A2:G' + backupSheet.getLastRow()).getValues();

  var restoredData = [];

  for (var i = 0; i < mainTitles.length; i++) {
    var title = mainTitles[i];
    var row = backupData.find(row => row[0] === title);
    if (row) {
      restoredData.push(row.slice(1)); // Columns B to G
    } else {
      restoredData.push(['', '', '', '', '', '']);
    }
  }

  mainSheet.getRange(2, 2, restoredData.length, restoredData[0].length).setValues(restoredData);
}
6. 🚨 Highlight Missing Rows in Backup Sheet
Purpose:
Automatically highlights rows in the backup sheet where the title no longer exists in the main sheet — allowing manual cleanup decisions.

How it works:
Compares Column A titles from main sheet and backup sheet.

If a title in backup is missing from main, highlights that row red.

Clears highlight if title reappears.

Script snippet:
javascript
Copy
Edit
function highlightMissingRowsInBackup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inboundSheet = ss.getSheetByName('THE INBOUND');
  const backupSheet = ss.getSheetByName('Backup');

  const inboundTitles = inboundSheet.getRange('A2:A' + inboundSheet.getLastRow()).getValues().flat().filter(String);

  const lastRow = backupSheet.getLastRow();
  const lastCol = backupSheet.getLastColumn();
  const backupRange = backupSheet.getRange(2, 1, lastRow - 1, lastCol);
  const backupData = backupRange.getValues();

  for (let i = 0; i < backupData.length; i++) {
    const title = backupData[i][0];
    const rowRange = backupSheet.getRange(i + 2, 1, 1, lastCol);

    if (title && !inboundTitles.includes(title)) {
      rowRange.setBackground('red');
    } else {
      rowRange.setBackground(null);
    }
  }
}
7. ⚙️ Find Missing Titles and Auto-Add to Backup
Purpose:
Checks for titles present in the main sheet but missing in backup — then offers to auto-add them.

How it works:
Reads Column A titles from main and backup sheets.

Filters for missing titles.

Prompts user with a dialog listing missing titles.

On confirmation, adds missing titles to backup sheet.

Script snippet:
javascript
Copy
Edit
function findMissingTitles() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const sheet1 = ss.getSheetByName("THE INBOUND");
  const sheet2 = ss.getSheetByName("Backup");

  if (activeSheet.getName() !== "Backup") {
    ui.alert("Please switch to the 'Backup' sheet to run this check.");
    return;
  }

  const titles1 = sheet1.getRange("A2:A" + sheet1.getLastRow()).getValues().flat().filter(String);
  const titles2 = sheet2.getRange("A2:A" + sheet2.getLastRow()).getValues().flat().filter(String);

  const missing = titles1.filter(title => !titles2.includes(title));

  if (missing.length > 0) {
    const message = "The following titles are missing from 'Backup':\n\n" + missing.join("\n") +
                    "\n\nWould you like me to add them for you automatically?";
    const response = ui.alert("Missing Titles Found", message, ui.ButtonSet.YES_NO);

    if (response === ui.Button.YES) {
      const startRow = sheet2.getLastRow() + 1;
      const range = sheet2.getRange("A" + startRow + ":A" + (startRow + missing.length - 1));
      const valuesToAdd = missing.map(title => [title]);
      range.setValues(valuesToAdd);
      ui.alert("✅ Missing titles have been added to the 'Backup' sheet.");
    } else {
      ui.alert("No titles were added.");
    }
  } else {
    ui.alert("✅ All titles from THE INBOUND are present in Backup.");
  }
}
💡 Need help adding buttons?
Each of these scripts can be bound to buttons or menu items for easy use.

For navigation buttons (e.g., “Accounts” or “Contacts”), make sure the active cell selection is passed as parameter.

I can provide button setup instructions if needed.

Final Notes
All scripts are written in Google Apps Script and tied to specific sheets by name.

Change sheet names and IDs as needed for your environment.

Use triggers (time-driven or on-edit) for automations like email sending or stage updates.

For easier navigation, make sure sheets have unique headers and consistent column indexes.
