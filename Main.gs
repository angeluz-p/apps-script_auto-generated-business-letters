// Replace with your folder IDs
const TEMPLATE_FOLDER_ID = "1V_ICjyNJ8h5RgoX77ba5j7x_70tbwGls";
const OUTPUT_FOLDER_ID = "17mKjRtv0kDRXzt5PkE2G7EN0L9-vDppn";
const INTERN_SUBFOLDER_NAME = "Interns";
const APPLICANT_SUBFOLDER_NAME = "Applicants";

// Sheet name configuration - change this once to update everywhere
const EMPLOYEE_SHEET_NAME = "Employee List";
const INTERN_SHEET_NAME = "Intern List";
const APPLICANT_SHEET_NAME = "Applicant List";

/**
 * @OnlyCurrentDoc false
 */

// Main function to show the letter generator modal
function showLetterGeneratorModal() {
  try {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('LetterGenerator')
      .setWidth(500)
      .setHeight(600)
      .setTitle('HR Letter Generator');
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Generate Letter');
  } catch (error) {
    console.error('Error showing modal:', error);
    SpreadsheetApp.getUi().alert(`Error: ${error.message}`);
  }
}

/** GET EVALUATION DATES */
// Core function to generate evaluation dates based on date range
function generateEvaluationDates(officialEndDate, extensionDate) {
  try {
    const startDate = new Date(officialEndDate);
    const endDate = new Date(extensionDate);
    
    if (startDate >= endDate) {
      console.error("Extension date must be after official end date");
      return [];
    }
    
    const evaluationDates = [];
    const currentDate = new Date(startDate);
    
    // Generate bimonthly evaluation dates
    while (currentDate <= endDate) {
      const monthYear = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "MMMM yyyy");
      evaluationDates.push(monthYear);
      
      // Move to next 2 months (bimonthly evaluation)
      currentDate.setMonth(currentDate.getMonth() + 2);
    }
    
    // Ensure the extension date month is included if not already present
    const extensionMonthYear = Utilities.formatDate(endDate, Session.getScriptTimeZone(), "MMMM yyyy");
    if (evaluationDates.length > 0 && evaluationDates[evaluationDates.length - 1] !== extensionMonthYear) {
      evaluationDates.push(extensionMonthYear);
    }
    
    return evaluationDates;
    
  } catch (error) {
    console.error("Error generating evaluation dates:", error);
    return [];
  }
}

// Generate formatted bullet points ready for template insertion
function generateEvaluationBulletPoints(officialEndDate, extensionDate) {
  try {
    const dates = generateEvaluationDates(officialEndDate, extensionDate);
    
    if (dates.length === 0) {
      return "";
    }
    
    let bulletPoints = "";
    
    for (let i = 0; i < dates.length; i++) {
      const date = dates[i];
      
      if (i === dates.length - 1) {
        // Last item gets special text
        bulletPoints += `     ●   ${date} (final evaluation before the end of the probationary period, or before the extended period expires)`;
      } else {
        bulletPoints += `     ●   ${date}`;
      }
      
      // Add newline if not the last item
      if (i < dates.length - 1) {
        bulletPoints += "\n";
      }
    }
    
    return bulletPoints;
    
  } catch (error) {
    console.error("Error generating bullet points:", error);
    return "";
  }
}

/** GET LETTER TYPE LIST FROM GOOGLE DRIVE */
function getTemplateByName(name) {
  try {
    const folder = DriveApp.getFolderById(TEMPLATE_FOLDER_ID);
    const files = folder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      try {
        const fileName = file.getName();
        if (fileName.toLowerCase().includes(name.toLowerCase())) {
          const mimeType = file.getMimeType();
          
          // If it's a shortcut, resolve it to the actual target file
          if (mimeType === "application/vnd.google-apps.shortcut") {
            try {
              const shortcutTargetId = file.getTargetId();
              if (shortcutTargetId) {
                const targetFile = DriveApp.getFileById(shortcutTargetId);
                console.log(`Resolved shortcut '${fileName}' to target: ${targetFile.getName()}`);
                return targetFile;
              } else {
                console.log(`Shortcut '${fileName}' has no valid target`);
                continue;
              }
            } catch (shortcutError) {
              console.log(`Error resolving shortcut '${fileName}': ${shortcutError.message}`);
              continue;
            }
          } else {
            // Return any file type (documents, spreadsheets, etc.)
            return file;
          }
        }
      } catch (permissionError) {
        console.log(`Skipped file due to permissions: ${permissionError.message}`);
      }
    }
    return null;
  } catch (error) {
    console.error("Error accessing template folder:", error);
    return null;
  }
}

// Get all template names from Google Drive
function getAllTemplateNames() {
  try {
    const folder = DriveApp.getFolderById(TEMPLATE_FOLDER_ID);
    const files = folder.getFiles();
    const templateNames = [];
    
    // List of letter types to hide (red highlighted items)
    const hiddenLetterTypes = [
      "AR for Retention Bonus",
      "Holiday Memo", 
      "Endorsement Letter for EW bank"
    ];
    
    while (files.hasNext()) {
      const file = files.next();
      
      try {
        const fileName = file.getName();
        const mimeType = file.getMimeType();
        
        // Include documents, shortcuts, AND Google Sheets
        if (
          mimeType === "application/vnd.google-apps.document" ||
          mimeType === "application/vnd.openxmlformats-officedocument.wordprocessingml.document" ||
          mimeType === "application/msword" ||
          mimeType === "application/vnd.google-apps.shortcut" ||
          mimeType === "application/vnd.google-apps.spreadsheet"
        ) {
          // Clean up the file name for dropdown display
          let cleanName = fileName
            .replace(/^Template - /i, "")
            .replace(/^Template/i, "")
            .replace(/ - Template\.docx$/i, "")
            .replace(/\.docx$/i, "")
            .replace(/^[\s\-]+|[\s\-]+$/g, "");
          
          // Check if this template should be hidden
          const shouldHide = hiddenLetterTypes.some(hiddenType => 
            cleanName.toLowerCase().includes(hiddenType.toLowerCase()) ||
            hiddenType.toLowerCase().includes(cleanName.toLowerCase())
          );
          
          // Only add to list if it's not in the hidden list
          if (!shouldHide) {
            templateNames.push({
              name: cleanName,
              fileName: fileName
            });
          }
        }
      } catch (permissionError) {
        console.log(`Permission denied for file: ${permissionError.message}`);
      }
    }
    
    return templateNames.sort((a, b) => a.name.localeCompare(b.name));
  } catch (error) {
    console.error("Error getting template names:", error);
    return [];
  }
}

/** GET LIST FROM EMPLOYEES, INTERNS AND APPLICANTS SHEET TABS */
// Get all people from Employee List, categorized by Column I
function getAllEmployees() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EMPLOYEE_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return [];
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    
    // Find relevant columns
    const firstNameCol = headers.findIndex(h => h.toString().toLowerCase().includes('first'));
    const lastNameCol = headers.findIndex(h => h.toString().toLowerCase().includes('last'));
    const emailCol = headers.findIndex(h => h.toString().toLowerCase().includes('email'));
    const jobTitleCol = headers.findIndex(h => h.toString().toLowerCase().includes('title'));
    const categoryCol = 8; // Column I is index 8 (0-based)
    
    const employees = data.map((row, index) => {
      const firstName = firstNameCol >= 0 ? row[firstNameCol] : '';
      const lastName = lastNameCol >= 0 ? row[lastNameCol] : '';
      const email = emailCol >= 0 ? row[emailCol] : '';
      const jobTitle = jobTitleCol >= 0 ? row[jobTitleCol] : '';
      const category = row[categoryCol] ? row[categoryCol].toString().toLowerCase().trim() : '';
      
      const fullName = `${firstName} ${lastName}`.trim();
      const displayName = fullName + (jobTitle ? ` (${jobTitle})` : '');
      
      return {
        rowIndex: index + 2,
        fullName: fullName,
        displayName: displayName,
        email: email,
        jobTitle: jobTitle,
        firstName: firstName,
        lastName: lastName,
        category: category
      };
    }).filter(emp => emp.fullName && emp.category !== 'applicant' && emp.category !== 'intern'); // Fixed logic
    
    return employees;
  } catch (error) {
    console.error('Error getting employees:', error);
    return [];
  }
}

// Get all interns from Employee List where Column I = "Intern"
function getAllInterns() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EMPLOYEE_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return [];
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    
    // Find relevant columns
    const firstNameCol = headers.findIndex(h => h.toString().toLowerCase().includes('first'));
    const lastNameCol = headers.findIndex(h => h.toString().toLowerCase().includes('last'));
    const emailCol = headers.findIndex(h => h.toString().toLowerCase().includes('email'));
    const jobTitleCol = headers.findIndex(h => h.toString().toLowerCase().includes('title') || h.toString().toLowerCase().includes('department'));
    const categoryCol = 8; // Column I is index 8 (0-based)
    
    const interns = data.map((row, index) => {
      const firstName = firstNameCol >= 0 ? row[firstNameCol] : '';
      const lastName = lastNameCol >= 0 ? row[lastNameCol] : '';
      const email = emailCol >= 0 ? row[emailCol] : '';
      const jobTitle = jobTitleCol >= 0 ? row[jobTitleCol] : '';
      const category = row[categoryCol] ? row[categoryCol].toString().toLowerCase().trim() : '';
      
      const fullName = `${firstName} ${lastName}`.trim();
      const displayName = fullName + (jobTitle ? ` (${jobTitle})` : '');
      
      return {
        rowIndex: index + 2,
        fullName: fullName,
        displayName: displayName,
        email: email,
        jobTitle: jobTitle,
        firstName: firstName,
        lastName: lastName,
        category: category
      };
    }).filter(intern => intern.fullName && intern.category === 'intern');
    
    console.log(`Found ${interns.length} interns`);
    return interns;
  } catch (error) {
    console.error('Error getting interns:', error);
    return [];
  }
}

// Get all applicants from Employee List where Column I = "Applicant"
function getAllApplicants() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EMPLOYEE_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      return [];
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    
    // Find relevant columns
    const firstNameCol = headers.findIndex(h => h.toString().toLowerCase().includes('first'));
    const lastNameCol = headers.findIndex(h => h.toString().toLowerCase().includes('last'));
    const emailCol = headers.findIndex(h => h.toString().toLowerCase().includes('email'));
    const jobTitleCol = headers.findIndex(h => h.toString().toLowerCase().includes('title'));
    const categoryCol = 8; // Column I is index 8 (0-based)
    
    const applicants = data.map((row, index) => {
      const firstName = firstNameCol >= 0 ? row[firstNameCol] : '';
      const lastName = lastNameCol >= 0 ? row[lastNameCol] : '';
      const email = emailCol >= 0 ? row[emailCol] : '';
      const jobTitle = jobTitleCol >= 0 ? row[jobTitleCol] : '';
      const category = row[categoryCol] ? row[categoryCol].toString().toLowerCase().trim() : '';
      
      const fullName = `${firstName} ${lastName}`.trim();
      const displayName = fullName + (jobTitle ? ` (${jobTitle})` : '');
      
      return {
        rowIndex: index + 2,
        fullName: fullName,
        displayName: displayName,
        email: email,
        firstName: firstName,
        lastName: lastName,
        category: category
      };
    }).filter(applicant => applicant.fullName && applicant.category === 'applicant');
    
    console.log(`Found ${applicants.length} applicants`);
    return applicants;
  } catch (error) {
    console.error('Error getting applicants:', error);
    return [];
  }
}

/** OLD DATA SILENT FUNCTION */
// Silent version of your generation function (no UI alerts)
function generateLetterWithDataSilent(dataObject, templateFile) {
  try {
    const mimeType = templateFile.getMimeType();
    
    // Check if it's a spreadsheet
    if (mimeType === "application/vnd.google-apps.spreadsheet") {
      return generateSpreadsheetWithDataSilent(dataObject, templateFile);
    }
    
    // Handle document templates (existing code)
    const outputFolder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
    const letterType = dataObject["Letter Type"];
    const fullName = dataObject["Full Name"] || `${dataObject["First Name"]} ${dataObject["Last Name"]}`;
    
    // Add current date to filename
    const currentDate = new Date();
    const dateString = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    const copiedDoc = templateFile.makeCopy(`${fullName} - ${letterType} - ${dateString}`, outputFolder);
    Utilities.sleep(1000);
    
    const doc = DocumentApp.openById(copiedDoc.getId());
    const body = doc.getBody();
    let replacementsMade = 0;
    
    Object.entries(dataObject).forEach(([key, value]) => {
      const placeholders = [
        `{{${key.replace(/\s+/g, "")}}}`,     // Remove all spaces: {{CurrentDay}}
        `{{${key}}}`,                          // Keep spaces: {{Current Day}}
        `{${key.replace(/\s+/g, "")}}`,       // Single braces no spaces: {CurrentDay}
        `{{${key.replace(/\s+/g, "").toLowerCase()}}}`, // Lowercase: {{currentday}}
        `{{${key.toLowerCase()}}}`,            // Lowercase with spaces: {{current day}}
      ];
      
      placeholders.forEach((placeholder) => {
        if (body.findText(placeholder)) {
          body.replaceText(placeholder, value || "");
          replacementsMade++;
        }
      });
    });
    
    doc.saveAndClose();
    Utilities.sleep(500);
    
    return {
      success: true,
      replacementsMade: replacementsMade,
      documentName: copiedDoc.getName(),
      fileUrl: copiedDoc.getUrl(),
    };
  } catch (error) {
    console.error("Error in generation:", error);
    throw error;
  }
}

function generateSpreadsheetWithDataSilent(dataObject, templateFile) {
  try {
    const outputFolder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
    const letterType = dataObject["Letter Type"];
    const fullName = dataObject["Full Name"] || `${dataObject["First Name"]} ${dataObject["Last Name"]}`;
    
    // Add current date to filename
    const currentDate = new Date();
    const dateString = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    const copiedFile = templateFile.makeCopy(`${fullName} - ${letterType} - ${dateString}`, outputFolder);
    Utilities.sleep(1000);
    
    const spreadsheet = SpreadsheetApp.openById(copiedFile.getId());
    const sheets = spreadsheet.getSheets();
    let replacementsMade = 0;
    
    // Process all sheets in the spreadsheet
    sheets.forEach(sheet => {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      
      if (lastRow > 0 && lastCol > 0) {
        const range = sheet.getRange(1, 1, lastRow, lastCol);
        const values = range.getValues();
        
        // Replace placeholders in each cell
        for (let row = 0; row < values.length; row++) {
          for (let col = 0; col < values[row].length; col++) {
            let cellValue = values[row][col];
            
            if (typeof cellValue === 'string' && cellValue.trim() !== '') {
              let originalValue = cellValue;
              
              Object.entries(dataObject).forEach(([key, value]) => {
                const placeholders = [
                  `{{${key.replace(/\s+/g, "")}}}`,     // Remove all spaces: {{CurrentDay}}
                  `{{${key}}}`,                          // Keep spaces: {{Current Day}}
                  `{${key.replace(/\s+/g, "")}}`,       // Single braces no spaces: {CurrentDay}
                  `{{${key.replace(/\s+/g, "").toLowerCase()}}}`, // Lowercase: {{currentday}}
                  `{{${key.toLowerCase()}}}`,            // Lowercase with spaces: {{current day}}
                ];
                
                placeholders.forEach(placeholder => {
                  if (cellValue.includes(placeholder)) {
                    cellValue = cellValue.replace(new RegExp(placeholder.replace(/[{}]/g, '\\$&'), 'g'), value || "");
                    replacementsMade++;
                  }
                });
              });
              
              // Update the cell value if it changed
              if (cellValue !== originalValue) {
                values[row][col] = cellValue;
              }
            }
          }
        }
        
        // Write back the updated values
        range.setValues(values);
      }
    });
    
    SpreadsheetApp.flush(); // Ensure all changes are saved
    Utilities.sleep(500);
    
    return {
      success: true,
      replacementsMade: replacementsMade,
      documentName: copiedFile.getName(),
      fileUrl: copiedFile.getUrl(),
    };
  } catch (error) {
    console.error("Error in spreadsheet generation:", error);
    throw error;
  }
}

/** NEW DATA SILENT FUNCTION */
// Function to get or create subfolder
function getOrCreateSubfolder(parentFolderId, subfolderName) {
  try {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const subfolders = parentFolder.getFoldersByName(subfolderName);
    
    if (subfolders.hasNext()) {
      // Subfolder exists, return it
      return subfolders.next();
    } else {
      // Subfolder doesn't exist, create it
      console.log(`Creating subfolder: ${subfolderName}`);
      return parentFolder.createFolder(subfolderName);
    }
  } catch (error) {
    console.error(`Error getting/creating subfolder ${subfolderName}:`, error);
    // Fallback to parent folder if subfolder creation fails
    return DriveApp.getFolderById(parentFolderId);
  }
}

// Generation function that accepts output folder parameter
function generateLetterWithDataSilentWithFolder(dataObject, templateFile, outputFolder) {
  try {
    const mimeType = templateFile.getMimeType();
    
    if (mimeType === "application/vnd.google-apps.spreadsheet") {
      return generateSpreadsheetWithDataSilentWithFolder(dataObject, templateFile, outputFolder);
    }
    
    const letterType = dataObject["Letter Type"];
    const fullName = dataObject["Full Name"] || `${dataObject["First Name"]} ${dataObject["Last Name"]}`;
    
    const currentDate = new Date();
    const dateString = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    const copiedDoc = templateFile.makeCopy(`${fullName} - ${letterType} - ${dateString}`, outputFolder);
    Utilities.sleep(1000);
    
    const doc = DocumentApp.openById(copiedDoc.getId());
    const body = doc.getBody();
    let replacementsMade = 0;
    
    // Handle QR Code insertion FIRST
    if (dataObject["QR Code URL"]) {
      try {
        const qrImageResponse = UrlFetchApp.fetch(dataObject["QR Code URL"]);
        const qrImageBlob = qrImageResponse.getBlob();
        
        // Replace text first, then find and replace with image
        body.replaceText("{{QRCode}}", "QR_IMAGE_HERE");
        
        const imageSearch = body.findText("QR_IMAGE_HERE");
        if (imageSearch) {
          const para = imageSearch.getElement().getParent().asParagraph();
          para.clear();
          para.appendInlineImage(qrImageBlob);
        }
        
      } catch (qrError) {
        body.replaceText("{{QRCode}}", "[QR Error]");
        body.replaceText("QR_IMAGE_HERE", "[QR Error]");
      }
    }
    
    // Handle regular text replacements
    Object.entries(dataObject).forEach(([key, value]) => {
      // Skip QR Code URL from text replacement
      if (key === "QR Code URL") return;
      
      const placeholders = [
        `{{${key.replace(/\s+/g, "")}}}`,
        `{{${key}}}`,
        `{${key.replace(/\s+/g, "")}}`,
        `{{${key.replace(/\s+/g, "").toLowerCase()}}}`,
        `{{${key.toLowerCase()}}}`,
      ];
      
      placeholders.forEach((placeholder) => {
        if (body.findText(placeholder)) {
          body.replaceText(placeholder, value || "");
          replacementsMade++;
        }
      });
    });
    
    doc.saveAndClose();
    Utilities.sleep(500);
    
    return {
      success: true,
      replacementsMade: replacementsMade,
      documentName: copiedDoc.getName(),
      fileUrl: copiedDoc.getUrl(),
      folderPath: outputFolder.getName()
    };
  } catch (error) {
    console.error("Error in generation:", error);
    throw error;
  }
}

// Spreadsheet generation function with folder parameter
function generateSpreadsheetWithDataSilentWithFolder(dataObject, templateFile, outputFolder) {
  try {
    const letterType = dataObject["Letter Type"];
    const fullName = dataObject["Full Name"] || `${dataObject["First Name"]} ${dataObject["Last Name"]}`;
    
    // Add current date to filename
    const currentDate = new Date();
    const dateString = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    const copiedFile = templateFile.makeCopy(`${fullName} - ${letterType} - ${dateString}`, outputFolder);
    Utilities.sleep(1000);
    
    const spreadsheet = SpreadsheetApp.openById(copiedFile.getId());
    const sheets = spreadsheet.getSheets();
    let replacementsMade = 0;
    
    // Process all sheets in the spreadsheet
    sheets.forEach(sheet => {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      
      if (lastRow > 0 && lastCol > 0) {
        const range = sheet.getRange(1, 1, lastRow, lastCol);
        const values = range.getValues();
        
        // Replace placeholders in each cell
        for (let row = 0; row < values.length; row++) {
          for (let col = 0; col < values[row].length; col++) {
            let cellValue = values[row][col];
            
            if (typeof cellValue === 'string' && cellValue.trim() !== '') {
              let originalValue = cellValue;
              
              Object.entries(dataObject).forEach(([key, value]) => {
                const placeholders = [
                  `{{${key.replace(/\s+/g, "")}}}`,     // Remove all spaces: {{CurrentDay}}
                  `{{${key}}}`,                          // Keep spaces: {{Current Day}}
                  `{${key.replace(/\s+/g, "")}}`,       // Single braces no spaces: {CurrentDay}
                  `{{${key.replace(/\s+/g, "").toLowerCase()}}}`, // Lowercase: {{currentday}}
                  `{{${key.toLowerCase()}}}`,            // Lowercase with spaces: {{current day}}
                ];
                
                placeholders.forEach(placeholder => {
                  if (cellValue.includes(placeholder)) {
                    cellValue = cellValue.replace(new RegExp(placeholder.replace(/[{}]/g, '\\$&'), 'g'), value || "");
                    replacementsMade++;
                  }
                });
              });
              
              // Update the cell value if it changed
              if (cellValue !== originalValue) {
                values[row][col] = cellValue;
              }
            }
          }
        }
        
        // Write back the updated values
        range.setValues(values);
      }
    });
    
    SpreadsheetApp.flush(); // Ensure all changes are saved
    Utilities.sleep(500);
    
    return {
      success: true,
      replacementsMade: replacementsMade,
      documentName: copiedFile.getName(),
      fileUrl: copiedFile.getUrl(),
      folderPath: outputFolder.getName()
    };
  } catch (error) {
    console.error("Error in spreadsheet generation:", error);
    throw error;
  }
}

/** MODAL FOR GENERATING LETTER */
// Generate letter from modal data (updated to use single sheet)
function generateLetterFromModal(employeeRowIndex, letterType, additionalFields = {}, isInternLetter = false, isApplicantLetter = false) {
  try {
    // Generate serial number
    const serialNumber = generateSerialNumber(letterType);
    console.log("Generated serial number:", serialNumber);
    
    // Create verification URL and QR code
    const verificationUrl = createVerificationURL(serialNumber);
    console.log("Verification URL:", verificationUrl);
    
    const qrCodeUrl = generateQRCode(verificationUrl);
    console.log("QR Code URL:", qrCodeUrl);
    
    // Add to additional fields
    additionalFields['Serial Number'] = serialNumber;
    additionalFields['QR Code URL'] = qrCodeUrl;

    if (letterType === 'Employee Incident Report') {
      console.log("Employee Incident Report detected - getting IR number");
      const currentIRNumber = getCurrentIRNumber();
      console.log("Current IR number:", currentIRNumber);
      additionalFields['Incident Number'] = currentIRNumber;
    }

    // Use single sheet name - always "Employee List"
    const sheetName = EMPLOYEE_SHEET_NAME;
    
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found. Available sheets: ${spreadsheet.getSheets().map(s => s.getName()).join(', ')}`);
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const employeeData = sheet.getRange(employeeRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const dataObject = processDataObjectWithMultipleDateFormats(headers, employeeData, additionalFields);
    dataObject["Letter Type"] = letterType;
    
    const templateFile = getTemplateByName(letterType);
    if (!templateFile) {
      throw new Error(`Template for '${letterType}' not found in folder.`);
    }
    
    let outputFolder;
    if (isApplicantLetter) {
      outputFolder = getOrCreateSubfolder(OUTPUT_FOLDER_ID, APPLICANT_SUBFOLDER_NAME);
    } else if (isInternLetter) {
      outputFolder = getOrCreateSubfolder(OUTPUT_FOLDER_ID, INTERN_SUBFOLDER_NAME);
    } else {
      outputFolder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
    }
    
    const result = generateLetterWithDataSilentWithFolder(dataObject, templateFile, outputFolder);

    if (letterType === 'Employee Incident Report' && result.success) {
      console.log("About to increment IR number for Employee Incident Report");
      incrementIRNumber();
      console.log("IR number increment completed");
    }
    
    return {
      success: true,
      message: `Letter generated successfully!`,
      documentName: result.documentName,
      fileUrl: result.fileUrl,
      replacementsMade: result.replacementsMade,
      currentDate: dataObject["Current Date"],
      folderPath: result.folderPath,
      serialNumber: serialNumber
    };
    
  } catch (error) {
    console.error('Error generating letter from modal:', error);
    throw new Error(`Generation failed: ${error.message}`);
  }
}

/** ADDITIONAL FIELDS */
function getAdditionalFieldsConfigWithDateFormats() {
  return {
    'Acknowledgment Receipt for Item Turnover': [
      { name: 'Turnover By', type: 'text', placeholder: 'Enter Name' },
      { name: 'Position', type: 'text', placeholder: 'Enter Position' },
    ],
    'Background Check': [
      { name: 'Recipient Name', type: 'text', placeholder: 'Enter Recipient Name' },
    ],
    'COE': [
      { name: 'Issued Place', type: 'text', placeholder: 'Enter Issued Place', value: 'Mandaluyong City' },
      // { name: 'Supervisor Email', type: 'email', placeholder: 'Enter Supervisor Email'},
      // { name: 'Human Resource Department', type: 'text', placeholder: 'Enter Name' },
      
    ],
    'Certificate of Completion for Intern': [
      { name: 'Hours', type: 'number', placeholder: 'Enter number of hours completed' },
      { name: 'Human Resource', type: 'text', placeholder: 'Enter HR Name' }
    ],
    'Employee Incident Report': [
      { name: 'Incident Number', type: 'text', placeholder: 'Loading...', readonly: true, autoPopulate: true },
      { name: 'Reported By', type: 'text', placeholder: 'Enter Name' },
      { name: 'Position', type: 'text', placeholder: 'Enter Position' },
      { name: 'Date Of Incident', type: 'date' },
      { name: 'Time Of Incident', type: 'time' },
      { name: 'Incident Location', type: 'text', placeholder: 'Enter Location' },
      { name: 'HR Representative', type: 'text', placeholder: 'Enter HR Representative Name' },
      { name: 'Operation Manager', type: 'text', placeholder: 'Enter Operation Manager Name' }
    ],
    'Extension for Probationary': [
      { name: 'Company Address', type: 'text', placeholder: 'Enter Company Address', value: 'Mandaluyong City' },
      { name: 'Management Name', type: 'text', placeholder: 'Enter Name' },
      { name: 'Official End Date', type: 'date' },
      { name: 'Extension Date', type: 'date' }
    ],
    'Item Acknowledgement Receipt': [
      { name: 'Issued By', type: 'text', placeholder: 'Enter Name' }
    ],
    'Job Offer':[
      { name: 'Employment Duration', type: 'text', placeholder: 'Enter Employment Duration (e.g., 6 months)' },
      { name: 'Key Responsibilities', type: 'richtext', placeholder: 'Enter key responsibilities and duties...' }
      //{ name: 'Report To', type: 'text', placeholder: 'Enter Name' },
      //{ name: 'Location', type: 'text', placeholder: 'Enter Location' }
    ],
    'Memo - Failure to Log in/out': [
      { name: 'From', type: 'text', placeholder: 'Enter From Name' },
      { name: 'Subject', type: 'text', placeholder: 'Enter Subject' },
      { name: 'Human Resource', type: 'text', placeholder: 'Enter HR Name' }
    ],
    'PAF': [
      { name: 'Cutoff', type: 'select', options: ['1st', '2nd'] },
      { name: 'Effective Date', type: 'date' },
      { name: 'Basic Salary Adjustment', type: 'number', placeholder: 'Enter Basic Salary Adjustment' },
      { name: 'Gross Pay From', type: 'number', placeholder: 'Enter Gross Pay Amount' },
      { name: 'Gross Pay To', type: 'number', placeholder: 'Enter Gross Pay Amount' },
      { name: 'Human Resource', type: 'text', placeholder: 'Enter HR Name' },
      { name: 'Approver', type: 'text', placeholder: 'Enter Approver Name' }
    ],
    'Prime - Promotion Announcement': [
      { name: 'Effective Date', type: 'date' }
    ],
    'Quit Claim and Waiver': [
      { name: 'Total Settlement Amount', type: 'text', placeholder: 'Enter Settlement Amount' },
      { name: 'Human Resource', type: 'text', placeholder: 'Enter HR Name' }
    ],
    'Undertaking and Agreement': [
      { name: 'Submission Deadline', type: 'date' },
      { name: 'Human Resource', type: 'text', placeholder: 'Enter HR Name' }
    ],
    'Vidalia - Promotion Announcement': [
      { name: 'Effective Date', type: 'date' }
    ]
  };
}

function getAdditionalFieldsConfig() {
  return getAdditionalFieldsConfigWithDateFormats();
}

/** FORMATTING FUNCTION */
// Creates multiple formats support
function processDataObjectWithMultipleDateFormats(headers, data, additionalFields = {}) {
  const dataObject = {};
  
  headers.forEach((header, index) => {
    const cleanKey = header.toString().trim();
    let value = data[index];

    if (value instanceof Date) {
      // Create multiple date formats for date fields
      const dateFormats = getMultipleDateFormats(value, cleanKey);
      Object.assign(dataObject, dateFormats);
    } else if (value !== null && value !== undefined) {
      if (cleanKey === "Basic Salary" || cleanKey === "Allowance" || cleanKey === "Total Settlement Amount") {
        value = formatCurrency(value);
      } else if (isNameField(cleanKey) && typeof value === 'string') {
        // Apply title case formatting to name fields
        value = toTitleCase(value.toString());
      } else if (isPositionField(cleanKey) && typeof value === 'string') {
        // Apply position title case formatting to position fields
        value = toPositionTitleCase(value.toString());
      } else if (isLocationField(cleanKey) && typeof value === 'string') {
        // Apply location title case formatting to location fields
        value = toLocationTitleCase(value.toString());
      } else {
        value = value.toString();
      }
      dataObject[cleanKey] = value;
    } else {
      dataObject[cleanKey] = "";
    }
  });

  // Handle additional fields with title case formatting
  Object.keys(additionalFields).forEach(key => {
    let value = additionalFields[key];
    
    // Check if this might be rich text content (contains HTML)
    if (value && typeof value === 'string' && value.includes('<')) {
      value = convertHtmlToPlainText(value);
    }
    
    if (value && isDateField(key)) {
      if (typeof value === 'string' && value.match(/^\d{4}-\d{2}-\d{2}$/)) {
        const dateObj = new Date(value + 'T00:00:00');
        const dateFormats = getMultipleDateFormats(dateObj, key);
        Object.assign(dataObject, dateFormats);
      }
    } else if (key === "Total Settlement Amount") {
      // Format Total Settlement Amount as currency
      dataObject[key] = formatCurrency(value);
    } else if (value && isNameField(key) && typeof value === 'string') {
      // Apply title case to name fields in additional fields
      dataObject[key] = toTitleCase(value);
    } else if (value && isPositionField(key) && typeof value === 'string') {
      // Apply position title case to position fields in additional fields
      dataObject[key] = toPositionTitleCase(value);
    } else if (value && isLocationField(key) && typeof value === 'string') {
      // Apply location title case to location fields in additional fields
      dataObject[key] = toLocationTitleCase(value);
    } else {
      dataObject[key] = value;
    }
  });

  // Add current date in multiple formats with enhanced ordinal support
  const currentDate = new Date();
  const currentDateFormats = getMultipleDateFormats(currentDate, "Current Date");
  Object.assign(dataObject, currentDateFormats);
  
  // Add specific current date components for common template usage
  const dayNumber = parseInt(Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "dd"));
  const ordinalDay = getOrdinalSuffix(dayNumber);
  
  dataObject["CurrentDay"] = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "dd");
  dataObject["CurrentDayOrdinal"] = ordinalDay;
  dataObject["CurrentMonth"] = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "MMMM");
  dataObject["CurrentYear"] = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy");

  // Enhanced gender-based title logic with title case formatting
  const gender = dataObject["Gender"] || dataObject["gender"] || "";
  const firstName = toTitleCase(dataObject["First Name"] || dataObject["FirstName"] || "");
  const lastName = toTitleCase(dataObject["Last Name"] || dataObject["LastName"] || "");
  const fullName = toTitleCase(dataObject["Full Name"] || dataObject["FullName"] || `${firstName} ${lastName}`.trim());
  
  // Generate title based on gender
  let title = "";
  const genderLower = gender.toString().toLowerCase().trim();
  
  if (genderLower === "male" || genderLower === "m") {
    title1 = "Mr.";
    title2Caps = "He";
    title2Low = "he";
    title3Caps = "His";
    title3Low = "his";
  } else if (genderLower === "female" || genderLower === "f") {
    title1 = "Ms.";
    title2Caps = "She";
    title2Low = "she";
    title3Caps = "Her";
    title3Low = "her";
  }

  // Update dataObject with title case formatted names
  dataObject["FirstName"] = firstName;
  dataObject["LastName"] = lastName;
  dataObject["FullName"] = fullName;
  
  dataObject["GenderTitle1"] = title1;
  dataObject["GenderTitle1FirstName"] = title1 ? `${title1} ${firstName}` : firstName;
  dataObject["GenderTitle1LastName"] = title1 ? `${title1} ${lastName}` : lastName;
  dataObject["GenderTitle1FullName"] = title1 ? `${title1} ${fullName}` : fullName;

  dataObject["GenderTitle2Caps"] = title2Caps;
  dataObject["GenderTitle2Low"] = title2Low;
  dataObject["GenderTitle3Caps"] = title3Caps;
  dataObject["GenderTitle3Low"] = title3Low;

  if (additionalFields['Official End Date'] && additionalFields['Extension Date']) {
    const officialEndDate = additionalFields['Official End Date'];
    const extensionDate = additionalFields['Extension Date'];
    
    // Generate evaluation dates
    const evaluationDates = generateEvaluationDates(officialEndDate, extensionDate);
    const bulletPoints = generateEvaluationBulletPoints(officialEndDate, extensionDate);
    
    // Add to data object with multiple format options
    dataObject['EvaluationBulletPoints'] = bulletPoints;
    dataObject['EvaluationDatesArray'] = evaluationDates;
    dataObject['EvaluationCount'] = evaluationDates.length;
    
    // Individual date placeholders for flexible template usage
    evaluationDates.forEach((date, index) => {
      dataObject[`Date${index + 1}`] = date;
      dataObject[`EvaluationDate${index + 1}`] = date;
    });
    
    // Fill remaining slots with empty strings (for templates expecting fixed positions)
    for (let i = evaluationDates.length; i < 6; i++) {
      dataObject[`Date${i + 1}`] = "";
      dataObject[`EvaluationDate${i + 1}`] = "";
    }
  }

  return dataObject;

}

function doGet(e) {
  const serialToVerify = e.parameter.verify;
  
  if (serialToVerify) {
    // You'll need to store verification data in a spreadsheet
    // For now, basic validation
    if (serialToVerify.match(/^(COE|BC|CCI)-\d{4}-\d{6}$/)) {
      return HtmlService.createHtmlOutput(`
        <html>
          <head>
            <style>
              body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
              .verified { color: #28a745; }
              .container { max-width: 500px; margin: 0 auto; }
            </style>
          </head>
          <body>
            <div class="container">
              <h2 class="verified">Certificate Verified</h2>
              <p><strong>Serial Number:</strong> ${serialToVerify}</p>
              <p><strong>Status:</strong> Valid</p>
              <p><strong>Verified:</strong> ${new Date().toLocaleDateString()}</p>
              <hr>
              <p><small>For additional verification, contact example@primeoutsourcing.com</small></p>
            </div>
          </body>
        </html>
      `);
    } else {
      return HtmlService.createHtmlOutput(`
        <html>
          <head>
            <style>
              body { font-family: Arial, sans-serif; text-align: center; padding: 50px; }
              .invalid { color: #dc3545; }
            </style>
          </head>
          <body>
            <div class="container">
              <h2 class="invalid">❌ Invalid Certificate</h2>
              <p>Serial number ${serialToVerify} is not valid.</p>
            </div>
          </body>
        </html>
      `);
    }
  }
  
  return HtmlService.createHtmlOutput(`
    <html>
      <body>
        <h2>Certificate Verification</h2>
        <p>Please scan a valid QR code.</p>
      </body>
    </html>
  `);
}