Excel VBA Macros for File Handling and Emailing
This project contains a set of VBA macros for automating the process of retrieving files from a specified folder, as well as sending emails with attachments using Microsoft Outlook.

Features:
1. Open Most Recent PDF
This macro opens the most recent PDF file in a specified folder, based on the folder path provided in the Parameters sheet (cell D11).
2. Send Email with Latest PDF Attachment
This macro opens a new email in Microsoft Outlook with the most recent PDF file from the specified folder attached.
The email recipients are dynamically fetched from D14 in the Parameters sheet and validated before sending.
3. Email Validation
The email addresses entered in D14 are validated to ensure they are in a correct email format.
Invalid email addresses will trigger an error message, preventing the email from being sent.
Setup Instructions:
Add Folder Path and Email Addresses:

Folder Path: In cell D11 of the Parameters sheet, enter the full path of the folder where the files are located (e.g., C:\Users\julie\Documents\Personal\BRO Example\Example 2025).
Email Recipients: In cell D14 of the Parameters sheet, enter the email addresses of the recipients. Separate multiple addresses with a semicolon (e.g., example1@example.com; example2@example.com).
Run Macros:

To open the most recent PDF, run the OpenMostRecentPDF macro.
To send an email with the latest PDF attached, run the EmailMostRecentPDF macro.
Macro Functionality:
Open Most Recent PDF:
Path Retrieval: The macro retrieves the folder path from cell D11 in the Parameters sheet.
File Detection: It identifies the most recent PDF based on the last modified date.
Opening the File: Once the latest PDF is identified, it opens the file in its default application (e.g., Adobe Reader).
Send Email with Latest PDF:
Path and Recipients: The macro fetches the folder path and email recipients from the Parameters sheet.
File Detection: It finds the most recent PDF file in the specified folder.
Email Creation: Using Outlook, it creates a new email with the latest PDF as an attachment.
Email Validation: The email addresses in D14 are validated before the email is sent. Invalid email addresses will trigger an error message, preventing the email from being sent.
Email Validation Details:
The email addresses are validated to ensure that they follow the correct format (e.g., user@example.com).
Multiple addresses can be entered in D14, separated by semicolons (;).
The macro checks for common errors, such as missing @ symbols, invalid domain names, or extra spaces.
Testing Invalid Email Scenarios:
You can test the email validation by entering the following types of invalid email addresses into D14:

Missing @ symbol: example.com
Invalid domain: example@domain
Extra spaces: example @example.com
Multiple @ symbols: example@@example.com
Missing TLD: example@domain.x
Blank email field: (empty cell)
The macro will show an error message and stop the process if any invalid email address is detected.

Additional Notes:
The macros use Microsoft Outlook for sending emails, so Outlook must be installed and configured on the system.
Make sure macros are enabled in Excel for the code to execute properly.
The PDF files to be opened or attached must be located in the specified folder. If no PDF files are found, the macros will show a corresponding error message.
