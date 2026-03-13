function getProjectSummary() {
  try {
    console.log('Starting getProjectSummary...');
    
    // Get the main database spreadsheet
    const spreadsheet = getMainDatabase();
    console.log('Database connected successfully');
    
    // Get the Projects sheet
    const sheet = spreadsheet.getSheetByName('Projects');
    if (!sheet) {
      const errorResult = {
        success: false,
        error: "Sheet 'Projects' not found"
      };
      console.log(JSON.stringify(errorResult, null, 2));
      return errorResult;
    }
    console.log('Projects sheet found');
    
    // Get all project data (assuming status is in a specific column, e.g., column C)
    const lastRow = sheet.getLastRow();
    console.log('Total rows in sheet: ' + lastRow);
    
    if (lastRow <= 1) {
      const emptyResult = {
        success: true,
        data: {
          total_projects: 0,
          ontrack_projects: 0,
          atrisk_projects: 0,
          delayed_projects: 0
        }
      };
      console.log('No projects found');
      console.log(JSON.stringify(emptyResult, null, 2));
      return emptyResult;
    }
    
    // Get status column data (starting from row 2 to skip header)
    const statusRange = sheet.getRange(2, 9, lastRow - 1, 1); // Column I for status
    const statusValues = statusRange.getValues();
    console.log('Status values retrieved: ' + statusValues.length + ' rows');
    
    // Initialize counters
    let totalProjects = 0;
    let ontrackProjects = 0;
    let atriskProjects = 0;
    let delayedProjects = 0;
    
    // Count projects by status
    statusValues.forEach(function(row) {
      const status = String(row[0]).toLowerCase().trim();
      if (status !== '' && status !== 'undefined' && status !== 'null') {
        totalProjects++;
        
        if (status === 'on track' || status === 'ontrack') {
          ontrackProjects++;
        } else if (status === 'at risk' || status === 'atrisk') {
          atriskProjects++;
        } else if (status === 'delayed') {
          delayedProjects++;
        }
      }
    });
    
    // Build the result
    const result = {
      success: true,
      data: {
        total_projects: totalProjects,
        ontrack_projects: ontrackProjects,
        atrisk_projects: atriskProjects,
        delayed_projects: delayedProjects
      }
    };
    
    // Log and return the summary
    console.log('Project summary completed:');
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


//Test function to verify getProjectSummary output.
function testGetProjectSummary() {
  const result = getProjectSummary();
  console.log(JSON.stringify(result, null, 2));
  return result;
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
