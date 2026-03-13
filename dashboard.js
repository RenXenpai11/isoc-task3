// dashboard.js
// Dashboard Summary Tool
// Reads data from the active sheet and logs a summary report

function myFunction() {
  console.log("hello world");
}

function getDashboardSummary() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  var totalRows = data.length - 1; // Exclude header row
  var headers = data[0];

  console.log("===== DASHBOARD SUMMARY =====");
  console.log("Sheet Name: " + sheet.getName());
  console.log("Total Columns: " + headers.length);
  console.log("Total Data Rows: " + totalRows);
  console.log("Headers: " + headers.join(", "));
  console.log("=============================");

  return {
    sheetName: sheet.getName(),
    totalRows: totalRows,
    totalColumns: headers.length,
    headers: headers
  };
}

function logTimestamp() {
  var now = new Date();
  console.log("Dashboard last checked: " + now.toLocaleString());
}

function runDashboard() {
  logTimestamp();
  getDashboardSummary();
}