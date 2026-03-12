/**
 * Retrieves the current budget value from the Finance sheet.
 * @returns {number|object} The budget value or an error object if retrieval fails.
 */
function getBudget() {
  try {
    // Get the budget value from Finance sheet, cell B19
    const result = getCellValue('Finance', 'B19');
    
    // Validate the result
    if (result === null || result === undefined || result === '') {
      console.log('Warning: Budget value is empty or not set');
      return 0;
    }
    
    // Log the result
    console.log('Budget retrieved successfully:', result);
    
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
