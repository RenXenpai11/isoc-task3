// GET ALL GENERAL TASKS
function getGeneralTasks() {
  try {
    const spreadsheet = getMainDatabase();
    const sheet = spreadsheet.getSheetByName("General Task");

    if (!sheet) {
      return {
        success: false,
        error: "Sheet 'General Task' not found"
      };
    }

    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return {
        success: true,
        tasks: []
      };
    }

    const headers = data[0];

    const tasks = data.slice(1).map(row => {
      let task = {};
      headers.forEach((header, index) => {
        task[header] = row[index];
      });
      return task;
    });

    return {
      success: true,
      tasks: tasks
    };

  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}


// CREATE GENERAL TASK
function createGeneralTask(title, description, status, priority, assignee, due_date) {
  try {
    const spreadsheet = getMainDatabase();
    const sheet = spreadsheet.getSheetByName("General Task");

    if (!sheet) {
      return {
        success: false,
        error: "Sheet 'General Task' not found"
      };
    }

    const newTask = [
      new Date().getTime(), // id
      title,
      description,
      status,
      priority,
      assignee,
      due_date,
      new Date() // created_at
    ];

    sheet.appendRow(newTask);

    return {
      success: true,
      message: "Task created successfully",
      task: {
        id: newTask[0],
        title: title,
        description: description,
        status: status,
        priority: priority,
        assignee: assignee,
        due_date: due_date
      }
    };

  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}


// TEST FUNCTION - CREATE TASK
function testCreateGeneralTask() {

  const result = createGeneralTask(
    "Fix login bug",
    "Login button not working",
    "To Do",
    "High",
    "John",
    "2026-03-20"
  );

  console.log(JSON.stringify(result, null, 2));
}


// TEST FUNCTION - GET TASKS
function testGetGeneralTasks() {

  const result = getGeneralTasks();

  console.log(JSON.stringify(result, null, 2));
}