# 🚀 **Google Apps Script Projects Portfolio**

Welcome to my **Google Apps Script Portfolio** — a collection of efficient automations designed to **streamline Google Sheets workflows**. Each project is built to improve **productivity, accuracy, and ease of use**.

---

## 1. 🔍 **Accounts Navigator Button**

**Purpose:**  
Quickly navigate & highlight account details in the *Accounts* sheet from the *Sales* or *Contacts* sheets.

**How it works:**  
- Select a name (e.g., “Naomi”) in the *Sales* sheet (Column C).  
- Click the **Accounts** button.  
- The script searches the **002_Accounts** sheet for matching rows, highlights them in yellow, and opens the sheet in a new tab.  
- Reverse navigation to *Contacts* is also supported.

![Accounts Navigator Button](./images/accounts-navigator-button.png)  
[📜 **View Full Script**](./scripts/checkAccounts.gs)

---

## 2. 🔄 **Automatic Sales Stage Updater**

**Purpose:**  
Automatically sets the sales stage to **“Offered”** for new entries with an RLM-Number but no stage set.

**How it works:**  
- Scans the *014_Sales* sheet rows.  
- Checks if Column B (RLM-Number) is filled & Column G (Stage) is empty.  
- Sets the Stage to **Offered** if conditions are met.

![Sales Stage Updater](./images/sales-stage-updater.png)  
[📜 **View Full Script**](./scripts/updateStage.gs)

---

## 3. 📧 **Automated Email Notification on Deal Closure**

**Purpose:**  
Sends email notifications automatically to departments when a deal is marked as **“Closed”**.

**How it works:**  
- Scans *014_Sales* for rows with **Stage = Closed**.  
- Sends detailed email notifications to recipients.  
- Marks the deal as **Email Sent** with timestamp to avoid duplicates.

![Email Notification Script](./images/email-notification-script.png)  
[📜 **View Full Script**](./scripts/sendEmailOnDealClosure.gs)

---

## 4. 📅 **Sorted Popup Sheet for Imported Data**

**Purpose:**  
Imports data from another spreadsheet, sorts by date, and displays it in a clean popup sheet.

**How it works:**  
- Imports data via URL from a source sheet.  
- Sorts by date ascending or descending.  
- Displays sorted data in a dedicated **SortedPopup** sheet.

![Sorted Popup Sheet](./images/sorted-popup-sheet.png)  
[📜 **View Full Script**](./scripts/sortImportrangePopupSheet.gs)

---

## 5. 🔄 **Backup Restore for Manual Data Sync**

**Purpose:**  
Keeps manual data synced between main and backup sheets, even after sorting or row deletions.

**How it works:**  
- Matches titles from main sheet with backup sheet.  
- Restores manual data columns in exact main sheet order.

![Backup Restore Script](./images/backup-restore-sync.png)  
[📜 **View Full Script**](./scripts/restoreManualData.gs)

---

## 6. 🚨 **Highlight Missing Rows in Backup Sheet**

**Purpose:**  
Highlights rows missing in main sheet within the backup sheet for manual cleanup.

**How it works:**  
- Compares titles in both sheets.  
- Highlights missing rows red in the backup sheet.

![Highlight Missing Rows](./images/highlight-missing-rows.png)  
[📜 **View Full Script**](./scripts/highlightMissingRowsInBackup.gs)

---

## 7. ⚙️ **Find Missing Titles and Auto-Add to Backup**

**Purpose:**  
Finds missing titles in backup sheet and offers to auto-add them.

**How it works:**  
- Lists missing titles with a prompt.  
- Adds missing titles to backup on confirmation.

![Find Missing Titles Script](./images/find-missing-titles-auto-add.png)  
[📜 **View Full Script**](./scripts/findMissingTitles.gs)

---

💡 **Tip:** Use consistent sheet names and unique headers for smooth script operation.  
Feel free to reach out if you want help with **buttons**, **menus**, or **trigger setups**!


