/**
 * Test function to get project summary
 * Run this to see the JSON output in the logs
 */
function getProjects() {
  const result = getProjectSummary();
  Logger.log('Project Summary:');
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * Web app endpoint - returns JSON
 * Deploy as Web App to access via HTTP
 */
function doGet(e) {
  try {
    const result = getProjectSummary();
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    const errorResponse = {
      success: false,
      error: error.message
    };
    return ContentService.createTextOutput(JSON.stringify(errorResponse))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Setup database connection - run this once
 */
function initializeDatabase() {
  try {
    setMainDatabase('1bB7JezupjhmQRj0MoMg0Agtpxu3yluhX13KJPISqMPU');
    Logger.log('✓ Database initialized successfully!');
    Logger.log('✓ Spreadsheet ID: 1bB7JezupjhmQRj0MoMg0Agtpxu3yluhX13KJPISqMPU');
  } catch (error) {
    Logger.log('✗ Database initialization failed: ' + error.message);
  }
}
