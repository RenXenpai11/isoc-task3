
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

function getProjectSummaryPayload() {
  return {
    success: true,
    data: {
      total_projects: 1,
      ontrack_projects: 1,
      atrisk_projects: 0,
      delayed_projects: 0
    }
  };
}

function doGet() {
  const payload = getProjectSummaryPayload();
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function testProjectSummaryPayload() {
  const payload = getProjectSummaryPayload();
  Logger.log(JSON.stringify(payload, null, 2));
  return payload;
}

