/**
 * Utility functions for the Google Sheets to Jira integration
 */

/**
 * Gets all settings from script properties
 * 
 * @return {Object} The settings object or null if not configured
 */
function getJiraSettings() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const settingsJson = scriptProperties.getProperty('jiraSettings');
  
  if (!settingsJson) {
    return null;
  }
  
  try {
    return JSON.parse(settingsJson);
  } catch (error) {
    Logger.log(`Error parsing settings: ${error.message}`);
    return null;
  }
}

/**
 * Saves settings to script properties
 * 
 * @param {Object} settings - The settings object to save
 */
function saveJiraSettings(settings) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('jiraSettings', JSON.stringify(settings));
}

/**
 * Gets column mapping from script properties
 * 
 * @return {Object} The column mapping object or empty object if not configured
 */
function getColumnMapping() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const mappingJson = scriptProperties.getProperty('columnMapping');
  
  if (!mappingJson) {
    return {};
  }
  
  try {
    return JSON.parse(mappingJson);
  } catch (error) {
    Logger.log(`Error parsing column mapping: ${error.message}`);
    return {};
  }
}

/**
 * Saves column mapping to script properties
 * 
 * @param {Object} mapping - The mapping object to save
 */
function saveColumnMapping(mapping) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('columnMapping', JSON.stringify(mapping));
}

/**
 * Gets headings from the active sheet
 * 
 * @return {Array} Array of column headings
 */
function getSheetHeadings() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const headings = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headings.filter(heading => heading !== '');
}

/**
 * Gets the column index for a specific heading
 * 
 * @param {Array} headings - Array of headings
 * @param {String} targetHeading - The heading to find
 * @return {Number} The column index (0-based) or -1 if not found
 */
function getColumnIndexByHeading(headings, targetHeading) {
  return headings.findIndex(heading => heading === targetHeading);
}

/**
 * Gets the column letter from a 0-based index
 * 
 * @param {Number} index - 0-based column index
 * @return {String} The column letter (A, B, C, etc.)
 */
function columnIndexToLetter(index) {
  let temp, letter = '';
  index += 1; // Convert to 1-based
  
  while (index > 0) {
    temp = (index - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    index = (index - temp - 1) / 26;
  }
  
  return letter;
}

/**
 * Converts row and column to A1 notation
 * 
 * @param {Number} row - 1-based row index
 * @param {Number} column - 0-based column index
 * @return {String} Cell in A1 notation
 */
function getCellA1Notation(row, column) {
  return `${columnIndexToLetter(column)}${row}`;
}

/**
 * Formats a date for display in Jira
 * 
 * @param {Date} date - The date to format
 * @return {String} Formatted date string
 */
function formatDateForJira(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss.SSSZ");
}

/**
 * Parses CSV list into array
 * 
 * @param {String} csvString - Comma-separated string
 * @return {Array} Array of trimmed values
 */
function parseCSVToArray(csvString) {
  if (!csvString) {
    return [];
  }
  
  return csvString.split(',')
    .map(item => item.trim())
    .filter(item => item !== '');
}

/**
 * Shows a toast message in the spreadsheet
 * 
 * @param {String} message - The message to show
 * @param {String} title - Optional title
 * @param {Number} timeoutSeconds - Optional timeout in seconds
 */
function showToast(message, title = 'Jira Integration', timeoutSeconds = 5) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, title, timeoutSeconds);
}

/**
 * Validates that required settings are configured
 * 
 * @return {Boolean} Whether settings are valid
 */
function validateSettings() {
  const settings = getJiraSettings();
  
  if (!settings) {
    showToast('Please configure Jira settings first!');
    return false;
  }
  
  const requiredFields = ['jiraUrl', 'email', 'apiToken', 'projectKey'];
  for (const field of requiredFields) {
    if (!settings[field]) {
      showToast(`Missing required setting: ${field}`);
      return false;
    }
  }
  
  return true;
}

/**
 * Gets data from the selected rows
 * 
 * @param {Array} headings - Array of column headings
 * @param {Object} mapping - Column to field mapping
 * @param {Range} selectedRange - The selected range
 * @return {Array} Array of data objects
 */
function getDataFromSelectedRows(headings, mapping, selectedRange) {
  const startRow = selectedRange.getRow();
  const numRows = selectedRange.getNumRows();
  const sheet = selectedRange.getSheet();
  
  // If selection starts in the header row, skip it
  const dataStartRow = startRow === 1 ? 2 : startRow;
  const dataNumRows = startRow === 1 ? numRows - 1 : numRows;
  
  if (dataNumRows <= 0) {
    return [];
  }
  
  // Get all data at once for better performance
  const data = sheet.getRange(dataStartRow, 1, dataNumRows, headings.length).getValues();
  const result = [];
  
  for (const row of data) {
    const rowData = {};
    const customFields = {};
    
    // Process regular fields
    for (let i = 0; i < headings.length; i++) {
      const heading = headings[i];
      const value = row[i];
      
      if (value === '' || value === undefined || value === null) {
        continue;
      }
      
      const mappedField = mapping[heading];
      if (!mappedField) {
        continue;
      }
      
      if (mappedField.startsWith('customfield_')) {
        customFields[mappedField] = value;
      } else {
        rowData[mappedField] = value;
      }
    }
    
    // If there are custom fields, add them to the row data
    if (Object.keys(customFields).length > 0) {
      rowData.customFields = customFields;
    }
    
    // Additional processing for specific fields
    if (rowData.labels && typeof rowData.labels === 'string') {
      rowData.labels = parseCSVToArray(rowData.labels);
    }
    
    if (rowData.components && typeof rowData.components === 'string') {
      rowData.components = parseCSVToArray(rowData.components);
    }
    
    result.push(rowData);
  }
  
  return result;
}

/**
 * Finds the first empty row in a specific column
 * 
 * @param {Sheet} sheet - The sheet to search in
 * @param {Number} column - The column index (0-based)
 * @return {Number} The first empty row (1-based) or -1 if none found
 */
function findFirstEmptyRow(sheet, column) {
  const data = sheet.getRange(2, column + 1, sheet.getLastRow() - 1, 1).getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (!data[i][0]) {
      return i + 2; // +2 because we start at row 2 and array is 0-based
    }
  }
  
  return -1;
}

/**
 * Updates a cell with the specified value
 * 
 * @param {Sheet} sheet - The sheet containing the cell
 * @param {Number} row - The row index (1-based)
 * @param {Number} column - The column index (0-based)
 * @param {String} value - The value to set
 */
function updateCell(sheet, row, column, value) {
  sheet.getRange(row, column + 1).setValue(value);
}

/**
 * Creates an HTML template from a file
 * 
 * @param {String} filename - Name of the HTML file
 * @param {Object} data - Data to pass to the template
 * @return {HtmlOutput} The processed template
 */
function createHtmlTemplate(filename, data = {}) {
  const template = HtmlService.createTemplateFromFile(filename);
  
  // Set template variables
  for (const [key, value] of Object.entries(data)) {
    template[key] = value;
  }
  
  return template.evaluate()
    .setWidth(600)
    .setHeight(500)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/**
 * Validates email format
 * 
 * @param {String} email - Email to validate
 * @return {Boolean} Whether the email is valid
 */
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Validates URL format
 * 
 * @param {String} url - URL to validate
 * @return {Boolean} Whether the URL is valid
 */
function isValidUrl(url) {
  try {
    new URL(url);
    return true;
  } catch (error) {
    return false;
  }
} 