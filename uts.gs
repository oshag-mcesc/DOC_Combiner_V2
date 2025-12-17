// ============================================================================
// SPREADSHEET DATA FUNCTIONS
// ============================================================================

/**
 * Loads configuration from the Config sheet
 * 
 * @returns {Object} Configuration object with all settings
 */
function loadConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  
  if (!configSheet) {
    throw new Error('Config sheet not found. Please create a sheet named "Config".');
  }
  
  const data = configSheet.getDataRange().getValues();
  const config = {};
  
  // Parse key-value pairs (assumes Setting in col A, Value in col B)
  for (let i = 1; i < data.length; i++) { // Skip header row
    const setting = data[i][0];
    const value = data[i][1];
    
    if (setting && setting !== '') {
      // Convert setting name to camelCase for object property
      const key = setting.replace(/\s+/g, '');
      const camelKey = key.charAt(0).toLowerCase() + key.slice(1);
      config[camelKey] = value;
    }
  }
  
  // Set default batch size if not specified
  if (!config.batchSize) {
    config.batchSize = DEFAULT_BATCH_SIZE;
  }
  
  return config;
}

/**
 * Loads document data from the DocIDs sheet
 * 
 * @returns {Array<Object>} Array of document data objects
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
 * 
 * @param {number} rowNumber - The row number (1-indexed)
 * @param {string} status - Status text (e.g., '✓', 'ERROR')
 * @param {string} errorMessage - Error message if applicable
 * @param {number} batchNumber - The batch number that processed this doc
 */
function updateDocumentStatus(rowNumber, status, errorMessage, batchNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const docSheet = ss.getSheetByName('DocIDs');
  
  if (!docSheet) return;
  
  // Columns: C = Status, D = Error Message, E = Batch #
  docSheet.getRange(rowNumber, 3).setValue(status);
  docSheet.getRange(rowNumber, 4).setValue(errorMessage);
  docSheet.getRange(rowNumber, 5).setValue(batchNumber);
  
  // Apply color coding
  if (status === '✓') {
    docSheet.getRange(rowNumber, 3).setBackground('#d9ead3'); // Light green
  } else if (status === 'ERROR') {
    docSheet.getRange(rowNumber, 3).setBackground('#f4cccc'); // Light red
  }
}

/**
 * Validates that all required configuration values are present
 * 
 * @param {Object} config - Configuration object
 * @returns {boolean} True if valid, false otherwise
 */
function validateConfig(config) {
  const required = ['outputFolderId', 'tempFolderId', 'pdfName'];
  
  for (const field of required) {
    if (!config[field] || config[field] === '') {
      Logger.log(`Missing required config field: ${field}`);
      return false;
    }
  }
  
  // Validate folder IDs exist
  try {
    DriveApp.getFolderById(config.outputFolderId);
    DriveApp.getFolderById(config.tempFolderId);
  } catch (error) {
    Logger.log('Invalid folder ID: ' + error.toString());
    return false;
  }
  
  return true;
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Resets all status fields in preparation for a new run
 */
function resetProcessStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clear DocIDs sheet status columns
  const docSheet = ss.getSheetByName('DocIDs');
  if (docSheet) {
    const lastRow = docSheet.getLastRow();
    if (lastRow > 1) {
      // Clear Status, Error Message, and Batch # columns
      docSheet.getRange(2, 3, lastRow - 1, 3).clearContent();
      docSheet.getRange(2, 3, lastRow - 1, 3).setBackground(null);
    }
  }
  
  // Reset Config sheet process status fields
  updateConfigValue('Process Status', 'NOT STARTED');
  updateConfigValue('Current Batch', 0);
  updateConfigValue('Total Batches', 0);
  updateConfigValue('Errors Count', 0);
  updateConfigValue('Start Time', '');
  updateConfigValue('Completion Time', '');
  updateConfigValue('Last Updated', '');
  updateConfigValue('Error Message', '');
  updateConfigValue('Final PDF URL', '');
  updateConfigValue('Final PDF ID', '');
  
  Logger.log('Process status reset');
}

/**
 * Clears all time-based triggers created by this script
 * Useful for stopping a running process or cleaning up
 */
function clearAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;
  
  triggers.forEach(trigger => {
    // Only delete triggers created by this script (time-based triggers)
    if (trigger.getEventType() === ScriptApp.EventType.CLOCK) {
      ScriptApp.deleteTrigger(trigger);
      count++;
    }
  });
  
  Logger.log(`Cleared ${count} triggers`);
  
  if (count > 0) {
    SpreadsheetApp.getUi().alert('Triggers Cleared', 
      `Removed ${count} scheduled triggers.`, 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Deletes the trigger that called the current function
 * Used to clean up after scheduled batch processing
 */
function deleteCurrentTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  
  // Delete the first time-based trigger (which should be the one that just fired)
  for (const trigger of triggers) {
    if (trigger.getEventType() === ScriptApp.EventType.CLOCK) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('Deleted current trigger');
      break;
    }
  }
}
