function getMainDatabase() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const spreadsheetId = scriptProperties.getProperty('MAIN_DB_ID');
  
  if (!spreadsheetId) {
    throw new Error('MAIN_DB_ID script property is not set. Please configure it in Project Settings > Script Properties.');
  }
  
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    return spreadsheet;
  } catch (error) {
    throw new Error('Failed to open spreadsheet with ID: ' + spreadsheetId + '. Error: ' + error.message);
  }
}

const SPREADSHEET_ID = '1bB7JezupjhmQRj0MoMg0Agtpxu3yluhX13KJPISqMPU'; //put the spreadsheet id here

function setMainDatabase(spreadsheetId = SPREADSHEET_ID) {
  if (!spreadsheetId || typeof spreadsheetId !== 'string') {
    throw new Error('Invalid spreadsheet ID. Please provide a valid spreadsheet ID string.');
  }
  
  // Validate that the spreadsheet exists and is accessible
  try {
    SpreadsheetApp.openById(spreadsheetId);
  } catch (error) {
    throw new Error('Cannot access spreadsheet with ID: ' + spreadsheetId + '. Please check the ID and permissions. Error: ' + error.message);
  }
  
  // Set the script property
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('MAIN_DB_ID', spreadsheetId);
  
  Logger.log('MAIN_DB_ID has been set to: ' + spreadsheetId);
  return true;
}

