// ============================================================================
// SPREADSHEET DATA FUNCTIONS
// ============================================================================

/**
 * Loads configuration from the Config sheet
 * Converts Setting/Value pairs into a configuration object
 * 
 * @returns {Object} Configuration object with all settings in camelCase
 * @throws {Error} If Config sheet is not found
 */
function loadConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  
  if (!configSheet) {
    throw new Error('Config sheet not found. Please create a sheet named "Config".');
  }
  
  const data = configSheet.getDataRange().getValues();
  const config = {};
  
  // Parse key-value pairs (Setting in column A, Value in column B)
  for (let i = 1; i < data.length; i++) { // Skip header row
    const setting = data[i][0];
    const value = data[i][1];
    
    if (setting && setting !== '') {
      // Convert setting name to camelCase for object property
      // Example: "Output Folder ID" → "outputFolderId"
      const key = setting.replace(/\s+/g, '');
      const camelKey = key.charAt(0).toLowerCase() + key.slice(1);
      config[camelKey] = value;
    }
  }
  
  return config;
}

/**
 * Loads document data from the DocIDs sheet
 * Extracts student names and document IDs
 * 
 * @returns {Array<Object>} Array of {studentName, docId} objects
 * @throws {Error} If DocIDs sheet is not found
 */
function loadDocumentData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const docSheet = ss.getSheetByName('DocIDs');
  
  if (!docSheet) {
    throw new Error('DocIDs sheet not found. Please create a sheet named "DocIDs".');
  }
  
  const data = docSheet.getDataRange().getValues();
  const docData = [];
  
  // Parse document data (assumes headers in row 1)
  for (let i = 1; i < data.length; i++) { // Skip header row
    const studentName = data[i][0];
    const docId = data[i][1];
    
    // Only include rows with both name and ID
    if (studentName && docId && docId !== '') {
      docData.push({
        studentName: studentName,
        docId: docId.trim()
      });
    }
  }
  
  return docData;
}

/**
 * Updates a single value in the Config sheet
 * If the setting doesn't exist, it will be added to the end
 * 
 * @param {string} setting - The setting name to update
 * @param {*} value - The new value
 */
function updateConfigValue(setting, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  
  if (!configSheet) return;
  
  const data = configSheet.getDataRange().getValues();
  
  // Find the row with this setting
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === setting) {
      configSheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  
  // If setting not found, add it to the end
  const lastRow = configSheet.getLastRow();
  configSheet.getRange(lastRow + 1, 1, 1, 2).setValues([[setting, value]]);
}

/**
 * Updates the status of a document in the DocIDs sheet
 * Applies color coding: green for success, red for errors
 * 
 * @param {number} rowNumber - The row number (1-indexed, includes header)
 * @param {string} status - Status text (e.g., '✓', 'ERROR')
 * @param {string} errorMessage - Error message if applicable
 */
function updateDocumentStatus(rowNumber, status, errorMessage) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const docSheet = ss.getSheetByName('DocIDs');
  
  if (!docSheet) return;
  
  // Update Status column (C) and Error Message column (D)
  docSheet.getRange(rowNumber, 3).setValue(status);
  docSheet.getRange(rowNumber, 4).setValue(errorMessage);
  
  // Apply color coding for visual feedback
  if (status === '✓') {
    docSheet.getRange(rowNumber, 3).setBackground('#d9ead3'); // Light green
  } else if (status === 'ERROR') {
    docSheet.getRange(rowNumber, 3).setBackground('#f4cccc'); // Light red
  }
}

/**
 * Validates that all required configuration values are present and valid
 * 
 * @param {Object} config - Configuration object
 * @returns {boolean} True if valid, false otherwise
 */
function validateConfig(config) {
  const required = ['outputFolderId', 'pdfName'];
  
  // Check that all required fields are present and not empty
  for (const field of required) {
    if (!config[field] || config[field] === '') {
      Logger.log(`Missing required config field: ${field}`);
      return false;
    }
  }
  
  // Validate that output folder ID exists and is accessible
  try {
    DriveApp.getFolderById(config.outputFolderId);
  } catch (error) {
    Logger.log('Invalid output folder ID: ' + error.toString());
    return false;
  }
  
  return true;
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Resets all status fields in preparation for a new run
 * Clears document statuses and resets Config sheet tracking values
 */
function resetProcessStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clear DocIDs sheet status columns (Status, Error Message)
  const docSheet = ss.getSheetByName('DocIDs');
  if (docSheet) {
    const lastRow = docSheet.getLastRow();
    if (lastRow > 1) {
      // Clear columns C and D (Status and Error Message)
      docSheet.getRange(2, 3, lastRow - 1, 2).clearContent();
      docSheet.getRange(2, 3, lastRow - 1, 2).setBackground(null);
    }
  }
  
  // Reset Config sheet process status fields
  updateConfigValue('Process Status', 'NOT STARTED');
  updateConfigValue('Errors Count', 0);
  updateConfigValue('Start Time', '');
  updateConfigValue('Completion Time', '');
  updateConfigValue('Error Message', '');
  updateConfigValue('Final PDF URL', '');
  updateConfigValue('Final PDF ID', '');
  
  Logger.log('Process status reset');
  
  SpreadsheetApp.getUi().alert('Status Reset', 
    'All status fields have been cleared and reset.', 
    SpreadsheetApp.getUi().ButtonSet.OK);
}