
function getCellValue(sheetName, cellRange) {
  try {
    // Get the main database spreadsheet
    const spreadsheet = getMainDatabase();
    
    // Get the specified sheet
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      return {
        success: false,
        error: `Sheet '${sheetName}' not found`
      };
    }
    
    // Get the range
    const range = sheet.getRange(cellRange);
    
    // Get values
    const values = range.getValues();
    
    // Return JSON based on whether it's a single cell or range
    if (values.length === 1 && values[0].length === 1) {
      // Single cell
      const result = {
        success: true,
        sheetName: sheetName,
        cellRange: cellRange,
        value: values[0][0]
      };
      //console.log(result);
      return result;
    } else {
      // Range of cells
      const result = {
        success: true,
        sheetName: sheetName,
        cellRange: cellRange,
        values: values
      };
      console.log(result);
      return result;
    }
    
  } catch (error) {
    const errorResult = {
      success: false,
      error: error.message
    };
    console.log(errorResult);
    return errorResult;
  }
}


// Test function to get project summary
function getProjects() {
  const result = getProjectSummary();
  Logger.log('Project Summary:');
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

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

function initializeDatabase() {
  try {
    setMainDatabase('1bB7JezupjhmQRj0MoMg0Agtpxu3yluhX13KJPISqMPU');
    Logger.log('✓ Database initialized successfully!');
    Logger.log('✓ Spreadsheet ID: 1bB7JezupjhmQRj0MoMg0Agtpxu3yluhX13KJPISqMPU');
  } catch (error) {
    Logger.log('✗ Database initialization failed: ' + error.message);
  }
}
