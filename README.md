📧 Automated Sales Data Email System (Google Apps Script)
![Photo1](https://github.com/mehwikhan9755-debug/Automated-sales-report-google-apps-script/blob/main/Screenshot%202025-11-24%20121604.png)
![photo2](https://github.com/mehwikhan9755-debug/Automated-sales-report-google-apps-script/blob/main/Screenshot%202025-11-24%20121625.png)
![photo3](https://github.com/mehwikhan9755-debug/Automated-sales-report-google-apps-script/blob/main/Screenshot%202025-11-24%20121705.png)
![photo3](https://github.com/mehwikhan9755-debug/Automated-sales-report-google-apps-script/blob/main/Screenshot%202025-11-24%20121826.png)

This project automates the process of filtering sales data by Salesman, generating a separate Google Sheet for each individual, and sending them their data via email — all in one click.

✅ No PDF
✅ Sends Google Sheet link
✅ Creates a unique sheet per Salesman
✅ Deletes temporary sheets after sending
✅ Email subject: "Your sales data in this sheet"
✅ Built with Google Apps Script

📌 Features

Reads sales records from Sheet1

Reads Salesman names and Email IDs from Sheet2

Filters sales data based on Salesman

Creates a new Google Sheet for each Salesman

Shares the sheet automatically with the Salesman

Emails the sheet link

Removes temporary sheets from the main spreadsheet

Fully automated with a single function run

📂 Google Sheet Structure
Sheet1 (Sales Data)

Headers:

Sale Date

Region

State

City

Product

Quantity

Salesman

AGE

Price per Unit

Sales Amount

Sheet2 (Email Mapping)

Headers:

Salesman

Email ID

🚀 How It Works

For each salesman listed in Sheet2:

Matches their name in Sheet1

Filters only their records

Creates a dedicated Google Sheet named:
SalesData_<SalesmanName>

Shares it with their email

Sends an automated email with the sheet link

Deletes the temporary sheet created during processing

🛠 Installation & Setup

Open your Google Sheet

Go to: Extensions → Apps Script

Delete existing code (if any)

Paste the script below

Click Save

Run: sendSalesDataToSalesmen()

Grant permissions when prompted

✅ Google Apps Script Code
function sendSalesDataToSalesmen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Sheet1");
  const emailSheet = ss.getSheetByName("Sheet2");

  const data = dataSheet.getDataRange().getValues();
  const emailData = emailSheet.getDataRange().getValues();
  const header = data[0];

  for (let i = 1; i < emailData.length; i++) {
    const salesman = emailData[i][0];
    const email = emailData[i][1];

    if (!salesman || !email) continue;

    const filteredRows = data.filter((row, index) => {
      return index === 0 || row[6] === salesman;
    });

    if (filteredRows.length <= 1) continue;

    const tempSheet = ss.insertSheet("Temp_" + salesman);
    tempSheet.getRange(1, 1, filteredRows.length, header.length)
      .setValues(filteredRows);

    const newSS = SpreadsheetApp.create("SalesData_" + salesman);
    const targetSheet = newSS.getActiveSheet();
    targetSheet.getRange(1, 1, filteredRows.length, header.length)
      .setValues(filteredRows);

    newSS.addEditor(email);

    MailApp.sendEmail({
      to: email,
      subject: "Your sales data in this sheet",
      body:
        "Dear " + salesman + ",\n\n" +
        "Please find your sales data in the Google Sheet below:\n" +
        newSS.getUrl() + "\n\n" +
        "Regards,\n" +
        "Mehvish Khan"
    });

    ss.deleteSheet(tempSheet);
  }
}

⏱ Optional: Automation

You can schedule automatic sending:

Go to Triggers

Click Add Trigger

Choose:

Function: sendSalesDataToSalesmen

Event: Time-Driven

Frequency: Daily/Weekly/Monthly

🧩 Requirements

Google Account

Google Sheets access

Apps Script permissions for:

Email sending

Drive file creation

Sheet editing

👩‍💼 Author

Mehvish Khan
Automating business workflows using Google Sheets & Apps Script.

⭐ Support

If you like this project, don’t forget to:

✅ Star ⭐ the repository
✅ Fork 🍴 to customize
✅ Share with others
