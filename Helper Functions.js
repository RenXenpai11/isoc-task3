
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
  const result = getProjectSummary(); //calls the getProjectSummary function to retrieve project summary data
  console.log('Project Summary:');
  console.log(JSON.stringify(result, null, 2));
  return result;
}

function doGet(e) { //handles GET requests to the web app and returns project summary data in JSON format
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

// Function to get project list
function displayProjectList(sheetName = 'Projects', rangeAddress = null) { 
  try {
    console.log('Fetching project list from sheet: ' + sheetName);
    const result = getProjectList(sheetName, rangeAddress);
    console.log('Project List:');
    console.log(JSON.stringify(result, null, 2));
    return result;
  } catch (error) {
    const errorResult = {
      success: false,
      error: error.message
    };
    console.log('Error fetching project list:');
    console.log(JSON.stringify(errorResult, null, 2));
    return errorResult;
  }
}

// Function to get project list with details
function getProjectList(sheetName = 'Projects', rangeAddress = null) {
  try {
    console.log('Starting getProjectList from sheet: ' + sheetName);
    
    // Get the main database spreadsheet
    const spreadsheet = getMainDatabase();
    console.log('Database connected successfully');
    
    // Get the specified sheet
    const sheet = spreadsheet.getSheetByName(sheetName); 
    if (!sheet) {
      return {
        success: false,
        error: "Sheet '" + sheetName + "' not found"
      };
    }
    console.log('Sheet "' + sheetName + '" found');
    
    // Get project data
    let projectData;
    
    if (rangeAddress) {
      // Use custom range if provided
      console.log('Using custom range: ' + rangeAddress);
      projectData = sheet.getRange(rangeAddress).getValues();
    } else {
      // Get all data from A2 to E (columns) and all rows
      const lastRow = sheet.getLastRow();
      
      if (lastRow <= 1) {
        return {
          success: true,
          data: []
        };
      }
      
      // Get range G2:K{lastRow}
      projectData = sheet.getRange(2, 7, lastRow - 1, 5).getValues();
    }
    
    console.log('Project data retrieved: ' + projectData.length + ' rows');
    
    // Process data into array of project objects
    const projects = [];
    
    try { 
      projectData.forEach(function(row, index) {
        try {
          const projectName = String(row[0]).trim();
          const status = String(row[2]).trim();
          const projectLeader = String(row[3]).trim();
          const progress = parseInt(row[4]) || 0;
          
          // Only add non-empty projects
          if (projectName !== '' && projectName !== 'undefined') {
            projects.push({
              project_name: projectName,
              status: status || 'N/A',
              project_leader: projectLeader || 'N/A',
              progress: progress
            });
          }
        } catch (rowError) {
          console.log('Error processing row ' + (index + 2) + ': ' + rowError.message);
        }
      });
    } catch (forEachError) {
      console.log('Error in forEach loop: ' + forEachError.message);
    }
    
    // Build the result
    const result = {
      success: true,
      data: projects
    };
    
    console.log('Project list completed:');
    console.log(JSON.stringify(result, null, 2));
    return result;
    
  } catch (error) {
    const errorResult = {
      success: false,
      error: error.message
    };
    console.log('Error occurred:');
    console.log(JSON.stringify(errorResult, null, 2));
    return errorResult;
  }
}
