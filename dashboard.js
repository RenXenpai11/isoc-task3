// dashboard.js
// Dashboard Summary Tool
// Reads data from the active sheet and logs a summary report

function myFunction() {
  console.log("Hello World! Dashboard is ready.");
}

// Gets a full summary of the active sheet
function getDashboardSummary() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  var totalRows = data.length - 1; // Exclude header row
  var headers = data[0];
  var lastUpdated = new Date().toLocaleString();

  console.log("===== DASHBOARD SUMMARY =====");
  console.log("Sheet Name   : " + sheet.getName());
  console.log("Total Columns: " + headers.length);
  console.log("Total Rows   : " + totalRows);
  console.log("Headers      : " + headers.join(", "));
  console.log("Last Updated : " + lastUpdated);
  console.log("=============================");

  return {
    sheetName: sheet.getName(),
    totalRows: totalRows,
    totalColumns: headers.length,
    headers: headers,
    lastUpdated: lastUpdated
  };
}

// Counts how many rows have data in the first column
function countNonEmptyRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var count = 0;

  for (var i = 1; i < data.length; i++) {  // Start at 1 to skip header
    if (data[i][0] !== "") {
      count++;
    }
  }

  console.log("Non-empty rows: " + count);
  return count;
}

// Highlights the header row in yellow
function highlightHeaders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headerRange = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  headerRange.setBackground("#FFD966");  // Yellow
  console.log("Header row highlighted!");
}

// Logs the current timestamp
function logTimestamp() {
  var now = new Date();
  console.log("Dashboard last checked: " + now.toLocaleString());
}

// Main function — runs everything
function runDashboard() {
  logTimestamp();
  getDashboardSummary();
  countNonEmptyRows();
  highlightHeaders();
  console.log("Dashboard run complete!");
}