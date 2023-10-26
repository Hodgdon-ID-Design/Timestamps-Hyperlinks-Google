# Timestamps-Hyperlinks-Google

These scripts were employed as part of a project aimed at assisting our LMS administrators in monitoring the addition and removal of users from the LMS system. The objective was to streamline their daily tasks by allowing them to refer to a linked master sheet that would automatically update timestamps whenever a linked document was edited. These codes were integrated with other Google Sheet formulas to maximize efficiency.

---
## Update Time Stamps & retrieve hyperlinks from a column

This script is designed to scan a specific column (column B by default) in a Google Sheets spreadsheet for hyperlinks that point to Google Sheets documents. Specifically, this is for items that do not have the hyperlink written out, but it linked in the text. For each valid hyperlink found, it retrieves the last modification timestamp of the linked Google Sheets document and updates a corresponding column (column G by default) with the formatted timestamp in the same row. The benefit of this script is that it doesn't require you to create a column dedicated to retrieve hyperlinks, and keep you document organized. However, sometimes the links themselves can be unstable.

* It operates on the currently active Google Sheets spreadsheet.

* The script loops through each row in the spreadsheet, starting from the first row (i.e., row 1).

* For each row, it retrieves the content of the cell in the second column (column B) for that row. This cell is assumed to contain a hyperlink.

* It checks if the hyperlink URL matches the pattern "/spreadsheets/d/" using a regular expression. This pattern is commonly found in Google Sheets URLs. If no valid URL is found, the script logs a message and continues to the next row.

* If a valid URL is found, it extracts the unique file ID from the URL using another regular expression.

* Using the extracted file ID, it retrieves the corresponding Google Drive file using **`DriveApp.getFileById(fileId)`**.

* It then obtains the last modification timestamp of that Google Drive file using **`file.getLastUpdated()`**.

* The last modification timestamp is formatted in the **`"MM/dd/yyyy hh:mm:ss a"`** format (e.g., "10/26/2023 02:30:45 PM") using **`Utilities.formatDate()`**.

* Finally, the formatted timestamp is written to the seventh column (column G) of the same row in the Google Sheets spreadsheet.

## Update Times from a hyperlink already writen in a cell

This script is designed to scan a specific column (column H by default) in a Google Sheets spreadsheet for URLs that point to Google Sheets documents. However, in this case, the URL is already written in the cell, eliminating the need for the second step to check the hyperlink. This code can serve as a backup in case the first code encounters errors by [extracting the hyperlink](https://www.youtube.com/watch?v=zEUQtvquJe0) and posting it in the column to be read.

* It begins by obtaining a reference to the currently active Google Sheets spreadsheet using **`SpreadsheetApp.getActiveSpreadsheet()`** and the active sheet within that spreadsheet using **`ss.getActiveSheet()`**.

* The code then enters a loop that iterates through the rows of the active sheet. The loop variable i is initialized to 1, and the loop continues until it reaches the last row of data in the sheet, as determined by **`sheet.getLastRow()`**.

* For each row in the loop, it retrieves the content of the cell in the eighth column (column H) for that row using **`sheet.getRange(i, 8)`** and stores it in the url variable.

* It checks if the content of the cell (url) matches a specific pattern using a regular expression. The pattern \/spreadsheets\/d\/ is used to identify Google Sheets URLs. If the pattern is not found in the URL, the code logs a message indicating that the row is skipped and continues to the next row using continue.

* If the URL matches the pattern, the code attempts to extract the unique file ID from the URL using a regular expression. The extracted fileId is the part of the URL that uniquely identifies the Google Sheets document.

* Using the extracted file ID, the code retrieves the corresponding Google Drive file using **`DriveApp.getFileById(fileId)`**.

* It then obtains the last modification timestamp of that Google Drive file using file.getLastUpdated().

* The last modification timestamp is formatted in the **`"MM/dd/yyyy hh:mm:ss a"`** format (e.g., "10/26/2023 02:30:45 PM") using **`Utilities.formatDate()`**.

* Finally, the formatted timestamp is written to the seventh column (column G) of the same row in the Google Sheets spreadsheet using **`sheet.getRange(i, 7).setValue(formattedDate)`**.
