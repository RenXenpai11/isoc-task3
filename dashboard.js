/**
 * Function to get the user's name from the Dashboard sheet
 */
function getUserName() {
  try {
    const result = getCellValue('Dashboard', 'C2');
    
    if (result.success) {
      const output = {
        success: true,
        data: {
          user_name: result.value
        }
      };
      console.log(JSON.stringify(output, null, 2));
      return output;
    } else {
      console.log(JSON.stringify({ success: false, error: result.error }, null, 2));
      return { success: false, error: result.error };
    }
  } catch (error) {
    console.log(JSON.stringify({ success: false, error: error.message }, null, 2));
    return { success: false, error: error.message };
  }
}

/**
 * Function to get the user's role/department from the Dashboard sheet
 */
function getRole() {
  try {
    const result = getCellValue('Dashboard', 'D2');
    
    if (result.success) {
      const output = {
        success: true,
        data: {
          role: result.value
        }
      };
      console.log(JSON.stringify(output, null, 2));
      return output;
    } else {
      console.log(JSON.stringify({ success: false, error: result.error }, null, 2));
      return { success: false, error: result.error };
    }
  } catch (error) {
    console.log(JSON.stringify({ success: false, error: error.message }, null, 2));
    return { success: false, error: error.message };
  }
}

/**
 * Function to get the password requirement from the Dashboard sheet
 */
function getPasswordRequirement() {
  try {
    const result = getCellValue('Dashboard', 'E2');
    
    if (result.success) {
      const output = {
        success: true,
        data: {
          password_requirement: result.value
        }
      };
      console.log(JSON.stringify(output, null, 2));
      return output;
    } else {
      console.log(JSON.stringify({ success: false, error: result.error }, null, 2));
      return { success: false, error: result.error };
    }
  } catch (error) {
    console.log(JSON.stringify({ success: false, error: error.message }, null, 2));
    return { success: false, error: error.message };
  }
}

/**
 * Function to validate login credentials
 */
function validateLogin(username, password) {
  try {
    const usernameResult = getCellValue('Dashboard', 'A2');
    const passwordResult = getCellValue('Dashboard', 'B2');
    
    if (!usernameResult.success || !passwordResult.success) {
      const output = { success: false, error: 'Unable to fetch login credentials' };
      console.log(JSON.stringify(output, null, 2));
      return output;
    }
    
    const sheetUsername = usernameResult.value;
    const sheetPassword = passwordResult.value;
    
    if (username === sheetUsername && password === sheetPassword) {
      const output = {
        success: true,
        data: {
          login_status: 'Authenticated',
          redirect_page: 'Dashboard'
        }
      };
      console.log(JSON.stringify(output, null, 2));
      return output;
    } else {
      const output = {
        success: true,
        data: {
          login_status: 'Denied',
          redirect_page: 'Login'
        }
      };
      console.log(JSON.stringify(output, null, 2));
      return output;
    }
  } catch (error) {
    const output = { success: false, error: error.message };
    console.log(JSON.stringify(output, null, 2));
    return output;
  }
}

/**
 * Run this to test all functions
 */
function testAll() {
  console.log('========== getUserName() ==========');
  getUserName();
  
  console.log('\n========== getRole() ==========');
  getRole();
  
  console.log('\n========== getPasswordRequirement() ==========');
  getPasswordRequirement();
  
  console.log('\n========== validateLogin(ana.reyes, SecurePass123) ==========');
  validateLogin('ana.reyes', 'SecurePass123');
  
  console.log('\n========== validateLogin(ana.reyes, WrongPassword) ==========');
  validateLogin('ana.reyes', 'WrongPassword');
}