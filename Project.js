function getBudget() {
  try {
    // Get the budget value from Finance sheet, cell B19
    const result = getCellValue('Finance', 'B19');
    
    // Log the result
    console.log(result);
    
    // Return the result
    return result;
    
  } catch (error) {
    const errorResult = {
      success: false,
      error: error.message
    };
    console.log(errorResult);
    return errorResult;
  }
}
