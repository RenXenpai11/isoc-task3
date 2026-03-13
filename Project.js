function getProjectSummary() { // Retrieves project summary data from the Projects sheet
  try {
    console.log('Starting getProjectSummary...');
    const spreadsheet = getMainDatabase();
    const sheet = spreadsheet.getSheetByName('Projects');
    
    if (!sheet) {
      console.log("Sheet 'Projects' not found");
      return { success: false, error: "Sheet 'Projects' not found" };
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      console.log('No projects found');
      return {
        success: true,
        data: {
          total_projects: 0,
          ontrack_projects: 0,
          atrisk_projects: 0,
          delayed_projects: 0 
        }
      };
    }
    
    const statusValues = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
    let totalProjects = 0, ontrackProjects = 0, atriskProjects = 0, delayedProjects = 0;
    
    statusValues.forEach(function(row) {
      const status = String(row[0]).toLowerCase().trim();
      if (status && status !== 'undefined' && status !== 'null') {
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
    
    const result = {
      success: true,
      data: {
        total_projects: totalProjects,
        ontrack_projects: ontrackProjects,
        atrisk_projects: atriskProjects,
        delayed_projects: delayedProjects
      }
    };
    console.log(JSON.stringify(result, null, 2));
    return result;
    
  } catch (error) {
    console.log('Error: ' + error.message);
    return { success: false, error: error.message };
  }
}

// Function to get project list with details
function getProjectList(sheetName = 'Projects', rangeAddress = 'G:K') { 
  try {
    console.log('Fetching projects from ' + sheetName);
    const spreadsheet = getMainDatabase();
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      console.log("Sheet '" + sheetName + "' not found");
      return { success: false, error: "Sheet '" + sheetName + "' not found" };
    }
    
    const projectData = sheet.getRange(rangeAddress).getValues();
    const projects = [];
    
    projectData.forEach(function(row) {
      const projectName = String(row[0]).trim();
      if (projectName && projectName !== 'undefined') {
        projects.push({
          project_name: projectName,
          status: String(row[2]).trim() || 'N/A',
          project_leader: String(row[3]).trim() || 'N/A',
          progress: parseInt(row[4]) || 0
        });
      }
    });
    
    console.log(JSON.stringify({ success: true, data: projects }, null, 2));
    return { success: true, data: projects };
    
  } catch (error) {
    console.log('Error: ' + error.message);
    return { success: false, error: error.message };
  }
}


// Retrieves task details from the Project Task sheettt
function getTaskDetails(sheetName = 'Project Task', rowNumber = 2) {
  try {
    console.log('Fetching task from ' + sheetName + ' row ' + rowNumber);
    const spreadsheet = getMainDatabase();
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      console.log("Sheet '" + sheetName + "' not found");
      return { success: false, error: "Sheet '" + sheetName + "' not found" };
    }
    
    const taskData = sheet.getRange(rowNumber, 1, 1, 10).getValues()[0];
    
    if (!taskData || taskData[0] === '') {
      console.log('No task found at row ' + rowNumber);
      return { success: false, error: "No task found at row " + rowNumber };
    }
    
    const result = { 
      success: true,
      data: {
        task_name: String(taskData[0]).trim() || 'N/A',
        priority: String(taskData[1]).trim() || 'N/A',
        assigned: String(taskData[2]).trim() || 'N/A',
        date: {
          start_date: String(taskData[3]).trim() || 'N/A',
          end_date: String(taskData[4]).trim() || 'N/A'
        },
        description: String(taskData[5]).trim() || 'N/A',
        status: String(taskData[6]).trim() || 'N/A',
        progress: parseInt(taskData[7]) || 0,
        notes: String(taskData[8]).trim() || 'N/A'
      }
    };
    console.log(JSON.stringify(result, null, 2));
    return result;
    
  } catch (error) {
    console.log('Error: ' + error.message);
    return { success: false, error: error.message };
  }
}