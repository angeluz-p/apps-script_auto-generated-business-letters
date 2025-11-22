// Function to get current IR number from Placeholders and References sheet
function getCurrentIRNumber() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName("Placeholders and References");
    
    if (!sheet) {
      console.error("Placeholders and References sheet not found");
      return "001";
    }
    
    const currentValue = sheet.getRange("A2").getValue();
    
    // Handle existing value - convert to string and ensure 3-digit format
    if (currentValue) {
      const valueStr = currentValue.toString().trim();
      
      // If it's already a number, format with leading zeros
      if (!isNaN(parseInt(valueStr))) {
        return parseInt(valueStr).toString().padStart(3, '0');
      }
      
      // If it's already formatted (like "003"), return as is
      if (valueStr.match(/^\d{3}$/)) {
        return valueStr;
      }
      
      // If it's some other format, try to extract number
      const numberMatch = valueStr.match(/\d+/);
      if (numberMatch) {
        return parseInt(numberMatch[0]).toString().padStart(3, '0');
      }
    }
    
    // Fallback - shouldn't reach here if there's already a value
    return "001";
    
  } catch (error) {
    console.error("Error getting IR number:", error);
    return "001";
  }
}

// Function to increment IR number after successful generation
function incrementIRNumber() {
  try {
    console.log("incrementIRNumber function called");
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName("Placeholders and References");
    
    if (!sheet) {
      console.error("Placeholders and References sheet not found");
      return;
    }
    
    const currentValue = sheet.getRange("A2").getValue();
    console.log("Current value in A2:", currentValue, "Type:", typeof currentValue);
    
    let currentNumber = 1;
    
    if (currentValue) {
      const valueStr = currentValue.toString().trim();
      console.log("Value as string:", valueStr);
      
      // Extract numeric value from current cell
      if (!isNaN(parseInt(valueStr))) {
        currentNumber = parseInt(valueStr);
        console.log("Parsed as number:", currentNumber);
      } else {
        // Try to extract number from string format
        const numberMatch = valueStr.match(/\d+/);
        if (numberMatch) {
          currentNumber = parseInt(numberMatch[0]);
          console.log("Extracted number from string:", currentNumber);
        }
      }
    }
    
    // Increment and update with 3-digit format
    const newNumber = currentNumber + 1;
    const newNumberStr = newNumber.toString().padStart(3, '0');
    console.log("Setting new value:", newNumberStr);
    
    sheet.getRange("A2").setValue(newNumberStr);
    
    // Force save
    SpreadsheetApp.flush();
    
    console.log(`IR number incremented from ${currentNumber.toString().padStart(3, '0')} to ${newNumberStr}`);
    
    // Verify the change was made
    const verifyValue = sheet.getRange("A2").getValue();
    console.log("Verification - value after update:", verifyValue);
    
  } catch (error) {
    console.error("Error incrementing IR number:", error);
    console.error("Error details:", error.stack);
  }
}