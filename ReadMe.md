Feature Set

-Google Sheets Version 

Feature Set

dashboard to provide overview of student status and financial summaries. 
Configuration sheet to provide variable settings
tabs for each student which are created automatically using their name. the sheet will contain student name, grade, instrument, ensemble list, parent name, parent cell, parent email, Amount Owed, Amount Paid, Amount Fundraised, Remaining Balance, and a list of financial transactions. Based off of the grade, instrument, and ensemble, each sheet will automatically create a financial transaction using the default values in the configuration sheet. The director can then add financial transactions as they occur, either payment, fundraise, or debt. 



I would like a way to integrate with gmail to be able to click a button on the student sheet and have it email the parent and student their current balance sheet. (this feature can be implemented later).





Dashboard Overview
Use JavaScript libraries like Chart.js or D3.js to create visual summaries of student statuses and financial data1.
Implement Google Charts for a more integrated approach with Google Sheets.
Importing Student Roster
Utilize the Google Sheets API to import data from a CSV file or another Google Sheet2.
Write a script to parse the CSV data and create individual tabs for each student.
Student Tab Structure
Each student tab can be dynamically created using JavaScript to include fields like name, grade, band participation, contact info, financial details, etc.
For the transaction area, use Google Sheets’ data validation feature to create dropdowns with pre-programmed costs.
Financial Tracking
Implement formulas in Google Sheets to calculate money owed, raised, and paid.
Use scripts to update these values automatically when new transactions are added.
Gmail Integration
To send emails directly from the sheet, you can use Google Apps Script to integrate with Gmail3.
Write a function that generates an email draft with the student’s balance sheet and sends it to both the student and the parent.
Here’s a basic JavaScript pseudocode outline to get you started:

// Function to import student roster and create tabs
function importStudentRoster() {
  // Code to import data from CSV or Google Sheet
  // Code to create a new tab for each student
}

// Function to add transaction with dropdown
function addTransaction(studentTab) {
  // Code to add a transaction to the student's tab
  // Include dropdown for pre-programmed costs and manual entry
}

// Function to calculate financial summaries
function calculateFinancialSummaries() {
  // Code to calculate money owed, raised, and paid
}

// Function to send email via Gmail
function sendEmail(parentEmail, studentEmail, balanceSheet) {
  // Code to draft and send an email with the balance sheet
}

// Call functions as needed
importStudentRoster();
// ... other function calls




