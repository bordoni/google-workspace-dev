/**
 * Google Sheets to Jira Ticket Integration
 * 
 * This script allows users to create Jira tickets from Google Sheets data.
 * It supports:
 * - Creating tickets based on sheet data
 * - Associating tickets with EPIC if specified in a cell
 * - Directing data to specific Jira fields based on column position
 */

// Add a custom menu when the spreadsheet opens
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Jira Integration')
    .addItem('Configure Settings', 'showSettingsDialog')
    .addItem('Configure Column Mapping', 'showColumnMappingDialog')
    .addSeparator()
    .addItem('Sync Ticket Status', 'syncJiraStatus')
    .addSeparator()
    .addItem('Test Connection', 'testJiraConnection')
    .addSeparator()
    .addItem('Install Edit Triggers', 'installEditTrigger')
    .addSeparator()
    .addSubMenu(ui.createMenu('Debug')
      .addItem('Show Debug Sheet', 'showDebugSheet')
      .addItem('Clear Debug Logs', 'clearDebugSheet')
      .addSeparator()
      .addItem('Analyze Missing Columns', 'analyzeSheetColumns')
      .addItem('View Column Mapping', 'viewColumnMapping')
      .addItem('Test Row Validation', 'testRowValidation')
      .addItem('Check Tab Settings', 'viewTabSettings'))
    .addToUi();
    
  // Show toast notification if Jira not configured
  if (!isJiraConfigured()) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Jira integration is not configured. Please use the "Configure Settings" option from the Jira Integration menu.',
      'Jira Setup Required',
      30
    );
  }
}

/**
 * Checks if required Jira settings are configured
 *
 * @since TBD
 * @return {boolean} True if all required settings are configured, false otherwise.
 */
function isJiraConfigured() {
  var settings = getSettings();
  return settings.jiraUrl && 
         settings.jiraEmail && 
         settings.jiraApiToken;
}

/**
 * Display a dialog for configuring Jira API settings
 */
function showSettingsDialog() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('Settings')
      .setWidth(600)
      .setHeight(500)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Jira Integration Settings');
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Jira Integration Settings');
  } catch (error) {
    Logger.log('Error showing settings dialog: ' + error);
    SpreadsheetApp.getUi().alert('Error displaying settings dialog: ' + error);
  }
}

/**
 * Display a dialog for configuring column mapping
 */
function showColumnMappingDialog() {
  var html = HtmlService.createHtmlOutputFromFile('ColumnMapping')
    .setWidth(600)
    .setHeight(500)
    .setTitle('Column Mapping Configuration');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Column Mapping Configuration');
}

/**
 * Save Jira settings to script properties
 */
function saveSettings(settings) {
  var scriptProperties = PropertiesService.getScriptProperties();
  
  // Save global settings
  scriptProperties.setProperties({
    'jiraUrl': settings.jiraUrl,
    'jiraEmail': settings.jiraEmail,
    'jiraApiToken': settings.jiraApiToken
  });
  
  // Save tab-specific settings
  if (settings.tabSettings) {
    scriptProperties.setProperty('tabSettings', JSON.stringify(settings.tabSettings));
  }
  
  return 'Settings saved successfully!';
}

/**
 * Get the current Jira settings
 *
 * @since TBD 
 * @return {Object} The settings object with default values if not set.
 */
function getSettings() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var settings = {
    jiraUrl: scriptProperties.getProperty('jiraUrl') || 'https://stellarwp.atlassian.net',
    jiraEmail: scriptProperties.getProperty('jiraEmail') || getCurrentUserEmail(),
    jiraApiToken: scriptProperties.getProperty('jiraApiToken'),
    tabSettings: {}
  };
  
  // Get tab-specific settings
  var tabSettingsStr = scriptProperties.getProperty('tabSettings');
  if (tabSettingsStr) {
    settings.tabSettings = JSON.parse(tabSettingsStr);
  }
  
  return settings;
}

/**
 * Get settings for a specific tab
 *
 * @since TBD
 * @param {string} tabName The name of the tab to get settings for.
 * @return {Object} Tab-specific settings with default values if not set.
 */
function getTabSettings(tabName) {
  var settings = getSettings();
  
  if (!tabName) {
    tabName = SpreadsheetApp.getActiveSheet().getName();
  }
  
  // Return tab-specific settings or defaults
  return settings.tabSettings[tabName] || {
    jiraProject: '',
    defaultIssueType: 'Task',
    epicFieldId: 'customfield_10014'
  };
}

/**
 * Gets the email address of the current user
 *
 * @since TBD
 * @return {string} The current user's email address.
 */
function getCurrentUserEmail() {
  try {
    return Session.getEffectiveUser().getEmail();
  } catch (error) {
    Logger.log('Error getting user email: ' + error);
    return '';
  }
}

/**
 * Save column mapping to script properties
 */
function saveColumnMapping(mapping) {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('columnMapping', JSON.stringify(mapping));
  return 'Column mapping saved successfully!';
}

/**
 * Get the current column mapping
 */
function getColumnMapping() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var mappingStr = scriptProperties.getProperty('columnMapping');
  
  if (mappingStr) {
    return JSON.parse(mappingStr);
  }
  
  return {};
}

/**
 * Get all column names from the active sheet
 */
function getSheetColumns() {
  var sheet = SpreadsheetApp.getActiveSheet();
  return sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
}

/**
 * Get all sheet names in the spreadsheet
 * 
 * @since TBD
 * @return {Array} Array of sheet names.
 */
function getAllSheetNames() {
  try {
    // Use the active spreadsheet that's already open in the context
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!spreadsheet) {
      Logger.log("No active spreadsheet found");
      return ["Sheet1"]; // Return a default sheet name
    }
    
    // Get all sheets and their names
    var sheets = spreadsheet.getSheets();
    var sheetNames = [];
    
    for (var i = 0; i < sheets.length; i++) {
      try {
        sheetNames.push(sheets[i].getName());
      } catch (e) {
        Logger.log("Error getting name for sheet at index " + i + ": " + e);
      }
    }
    
    Logger.log("Retrieved sheet names: " + JSON.stringify(sheetNames));
    
    // If we couldn't get any sheet names, return a default
    if (sheetNames.length === 0) {
      return ["Sheet1"];
    }
    
    return sheetNames;
  } catch (error) {
    Logger.log("Error getting sheet names: " + error);
    // Return a default sheet name instead of throwing an error
    return ["Sheet1"];
  }
}

/**
 * Test the Jira connection using saved credentials
 */
function testJiraConnection() {
  var settings = getSettings();
  
  if (!settings.jiraUrl || !settings.jiraEmail || !settings.jiraApiToken) {
    SpreadsheetApp.getUi().alert('Please configure Jira settings first!');
    return;
  }
  
  try {
    var response = makeJiraRequest('myself', 'GET');
    var responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      var userData = JSON.parse(response.getContentText());
      SpreadsheetApp.getUi().alert('Connection successful!\nLogged in as: ' + userData.displayName);
    } else {
      SpreadsheetApp.getUi().alert('Connection failed.\nResponse code: ' + responseCode);
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error connecting to Jira: ' + error.toString());
  }
}

/**
 * Make an API request to Jira
 */
function makeJiraRequest(endpoint, method, payload) {
  var settings = getSettings();
  
  if (!settings.jiraUrl || !settings.jiraEmail || !settings.jiraApiToken) {
    throw new Error('Jira settings not configured');
  }
  
  var url = settings.jiraUrl.replace(/\/$/, '') + '/rest/api/2/' + endpoint;
  var authHeader = 'Basic ' + Utilities.base64Encode(settings.jiraEmail + ':' + settings.jiraApiToken);
  
  var options = {
    'method': method,
    'headers': {
      'Authorization': authHeader,
      'Content-Type': 'application/json'
    },
    'muteHttpExceptions': true
  };
  
  if (payload) {
    options.payload = JSON.stringify(payload);
  }
  
  return UrlFetchApp.fetch(url, options);
}

/**
 * Updates a cell with a hyperlink to the Jira ticket
 *
 * @since TBD
 * @param {Sheet} sheet The sheet containing the cell
 * @param {number} row Row index (1-based)
 * @param {number} column Column index (1-based)
 * @param {string} ticketKey The Jira ticket key (e.g., PROJ-123)
 */
function updateCellWithTicketLink(sheet, row, column, ticketKey) {
  if (!ticketKey) return;
  
  var settings = getSettings();
  var ticketUrl = settings.jiraUrl.replace(/\/$/, '') + '/browse/' + ticketKey;
  var formula = '=HYPERLINK("' + ticketUrl + '", "' + ticketKey + '")';
  
  try {
    Logger.log("Setting hyperlink in cell [" + row + "," + column + "]: " + formula);
    sheet.getRange(row, column).setFormula(formula);
  } catch (error) {
    Logger.log("Error setting hyperlink: " + error);
  }
}

/**
 * Shows a toast notification in the spreadsheet
 *
 * @since TBD
 * @param {string} message The message to display
 * @param {string} title Optional title for the toast
 * @param {number} timeoutSeconds How long to display the toast (default: 5 seconds)
 */
function showToast(message, title, timeoutSeconds) {
  title = title || 'Jira Integration';
  timeoutSeconds = timeoutSeconds || 5;
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    message,
    title,
    timeoutSeconds
  );
}

/**
 * Create Jira tickets from selected sheet data
 */
function createJiraTicketsFromSelection() {
  if (!isJiraConfigured()) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      'Jira Settings Required',
      'Your Jira integration is not configured yet. Would you like to configure it now?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      showSettingsDialog();
    }
    return;
  }
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var selectedRange = SpreadsheetApp.getActiveRange();
  
  if (!selectedRange) {
    SpreadsheetApp.getUi().alert('Please select data rows to create tickets');
    return;
  }
  
  // Get column headers
  var headers = selectedRange.getRowIndex() === 2 ? selectedRange.getValues()[0] : sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  var dataRows = selectedRange.getValues();
  var statusColumn = findColumnByMapping(headers, 'status', 'Status');
  var ticketKeyColumn = findTicketKeyColumn(headers);
  var numberColumn = headers.indexOf('#') + 1;
  
  var createdCount = 0;
  var errorCount = 0;
  var errorMessages = [];
  var skippedCount = 0;
  var validationErrors = [];
  
  // First check all rows for validation errors
  for (var i = 0; i < dataRows.length; i++) {
    var rowIndex = selectedRange.getRowIndex() + i;
    
    // Skip rows that already have a ticket key
    if (ticketKeyColumn > 0 && sheet.getRange(rowIndex, ticketKeyColumn).getValue()) {
      skippedCount++;
      continue;
    }
    
    // Validate the row
    var validation = validateRowForTicket(sheet, rowIndex);
    
    if (!validation.success) {
      validationErrors.push({
        row: rowIndex,
        errors: validation.errors,
        missingFields: validation.missingRequiredFields
      });
    }
  }
  
  // If we have validation errors, show them and ask if the user wants to proceed with valid rows only
  if (validationErrors.length > 0) {
    var ui = SpreadsheetApp.getUi();
    var errorMessage = "The following rows have validation errors and cannot be processed:\n\n";
    
    validationErrors.forEach(function(error) {
      errorMessage += "Row " + error.row + ": " + error.errors.join(", ") + "\n";
    });
    
    errorMessage += "\nDo you want to proceed with creating tickets for valid rows only?";
    
    var response = ui.alert(
      'Validation Errors',
      errorMessage,
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return; // User cancelled
    }
  }
  
  // Process valid rows
  for (var i = 0; i < dataRows.length; i++) {
    var rowData = dataRows[i];
    var rowIndex = selectedRange.getRowIndex() + i;
    
    // Skip rows that already have a ticket key
    if (ticketKeyColumn > 0 && sheet.getRange(rowIndex, ticketKeyColumn).getValue()) {
      continue;
    }
    
    // Skip rows with validation errors
    var hasValidationError = validationErrors.some(function(error) {
      return error.row === rowIndex;
    });
    
    if (hasValidationError) {
      continue;
    }
    
    try {
      // Show toast that we're creating a ticket
      var rowDescription = "Row " + rowIndex;
      if (headers.indexOf('Summary') >= 0 && rowData[headers.indexOf('Summary')]) {
        rowDescription = rowData[headers.indexOf('Summary')].toString().substring(0, 30);
        if (rowData[headers.indexOf('Summary')].toString().length > 30) {
          rowDescription += "...";
        }
      }
      
      showToast("Creating Jira ticket for: " + rowDescription, "Creating Ticket", 3);
      
      var ticketData = prepareTicketDataWithMapping(headers, rowData);
      var response = createJiraTicket(ticketData);
      
      var responseData = JSON.parse(response.getContentText());
      var ticketKey = responseData.key;
      
      // Update the ticket key in the sheet
      if (ticketKeyColumn > 0) {
        sheet.getRange(rowIndex, ticketKeyColumn).setValue(ticketKey);
      }
      
      // Update the # column with a hyperlink to the ticket
      if (numberColumn > 0) {
        updateCellWithTicketLink(sheet, rowIndex, numberColumn, ticketKey);
      }
      
      // Update any columns mapped to ticketId
      if (ticketData._ticketIdColumns && ticketData._ticketIdColumns.length > 0) {
        ticketData._ticketIdColumns.forEach(function(columnInfo) {
          var columnIndex = columnInfo.index + 1; // Convert to 1-based index
          logToSheet('createJiraTicketsFromSelection', 'info', 'Updating ticketId column', {
            header: columnInfo.header,
            column: columnIndex,
            row: rowIndex,
            ticketKey: ticketKey
          });
          
          // Update the cell with a hyperlink to the ticket
          updateCellWithTicketLink(sheet, rowIndex, columnIndex, ticketKey);
        });
      }
      
      // Update the status if status column exists
      if (statusColumn > 0) {
        sheet.getRange(rowIndex, statusColumn).setValue('Created');
      }
      
      // Show success toast
      showToast("Successfully created ticket " + ticketKey, "Success", 5);
      
      createdCount++;
      
    } catch (error) {
      errorCount++;
      var errorMsg = error.toString();
      errorMessages.push('Error on row ' + rowIndex + ': ' + errorMsg);
      
      // Show error toast
      showToast("Error creating ticket: " + errorMsg.substring(0, 50), "Error", 8);
    }
  }
  
  var message = createdCount + ' ticket(s) created successfully.\n';
  
  if (skippedCount > 0) {
    message += skippedCount + ' row(s) skipped (already have tickets).\n';
  }
  
  if (validationErrors.length > 0) {
    message += validationErrors.length + ' row(s) skipped due to validation errors.\n';
  }
  
  if (errorCount > 0) {
    message += errorCount + ' error(s) occurred during creation:\n' + errorMessages.join('\n');
  }
  
  SpreadsheetApp.getUi().alert(message);
}

/**
 * Create Jira tickets from the entire sheet (bulk creation)
 */
function createJiraTicketsFromSheet() {
  if (!isJiraConfigured()) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      'Jira Settings Required',
      'Your Jira integration is not configured yet. Would you like to configure it now?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      showSettingsDialog();
    }
    return;
  }
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert('No data rows found in the sheet');
    return;
  }
  
  // Confirm with the user
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Confirm Bulk Creation',
    'This will create Jira tickets for all rows in the sheet that don\'t already have a ticket key. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  // Get column headers
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  
  // Get column indexes
  var statusColumn = findColumnByMapping(headers, 'status', 'Status');
  var ticketKeyColumn = findTicketKeyColumn(headers);
  var processColumn = headers.indexOf('Process') + 1;
  
  // Get all data rows
  var dataRows = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  
  var createdCount = 0;
  var skippedCount = 0;
  var errorCount = 0;
  var errorMessages = [];
  
  // Create a progress dialog
  var htmlOutput = HtmlService
    .createHtmlOutput('<p>Creating Jira tickets...</p>')
    .setWidth(300)
    .setHeight(100);
  var dialog = ui.showModelessDialog(htmlOutput, 'Progress');
  
  for (var i = 0; i < dataRows.length; i++) {
    var rowData = dataRows[i];
    var rowIndex = i + 2; // Add 2 to account for 1-based index and header row
    
    // Skip rows that already have a ticket key
    if (ticketKeyColumn > 0 && rowData[ticketKeyColumn - 1]) {
      skippedCount++;
      continue;
    }
    
    // Skip rows if process column exists and is not set to TRUE
    if (processColumn > 0 && rowData[processColumn - 1] !== true) {
      skippedCount++;
      continue;
    }
    
    try {
      var ticketData = prepareTicketDataWithMapping(headers, rowData);
      var response = createJiraTicket(ticketData);
      
      var responseData = JSON.parse(response.getContentText());
      
      // Update the ticket key in the sheet
      if (ticketKeyColumn > 0) {
        sheet.getRange(rowIndex, ticketKeyColumn).setValue(responseData.key);
      }
      
      // Update the status if status column exists
      if (statusColumn > 0) {
        sheet.getRange(rowIndex, statusColumn).setValue('Created');
      }
      
      createdCount++;
      
    } catch (error) {
      errorCount++;
      errorMessages.push('Error on row ' + rowIndex + ': ' + error.toString());
    }
    
    // Allow UI to refresh every 10 rows
    if (i % 10 === 0) {
      SpreadsheetApp.flush();
    }
  }
  
  // Close the progress dialog
  dialog.close();
  
  var message = createdCount + ' ticket(s) created successfully.\n' +
                skippedCount + ' row(s) skipped.\n';
  if (errorCount > 0) {
    message += errorCount + ' error(s) occurred:\n' + errorMessages.join('\n');
  }
  
  SpreadsheetApp.getUi().alert(message);
}

/**
 * Prepare ticket data from row data using the configured column mapping
 */
function prepareTicketDataWithMapping(headers, rowData) {
  var settings = getSettings();
  var tabSettings = getTabSettings();
  var columnMapping = getColumnMapping();
  var epicFieldId = getEpicFieldId();
  
  // Log ticket preparation start
  logToSheet('prepareTicketDataWithMapping', 'info', 'Preparing ticket data with mapping:');
  logToSheet('prepareTicketDataWithMapping', 'info', 'Headers', headers);
  logToSheet('prepareTicketDataWithMapping', 'info', 'Tab settings', tabSettings);
  logToSheet('prepareTicketDataWithMapping', 'info', 'Epic Field ID', epicFieldId);
  
  var ticketData = {
    fields: {
      issuetype: {
        name: tabSettings.defaultIssueType
      }
    }
  };
  
  // Track columns mapped to ticketId for later use when updating the sheet
  var ticketIdColumns = [];
  
  // Check for EPIC column without mapping
  var epicColumnIndex = findColumnByMapping(headers, 'epicLink', ['EPIC', 'Epic'], columnMapping) - 1; // Convert to 0-based
  
  Logger.log('Epic column index: ' + epicColumnIndex);
  
  var epicValue = null;
  var projectKey = null;
  
  // If EPIC column exists and has a value
  if (epicColumnIndex >= 0 && rowData[epicColumnIndex]) {
    epicValue = rowData[epicColumnIndex];
    ticketData.fields[epicFieldId] = epicValue;
    
    // Try to extract project key from EPIC (format typically "PROJECT-123")
    var epicParts = epicValue.split('-');
    if (epicParts.length > 1) {
      projectKey = epicParts[0];
      Logger.log("Extracted project key from EPIC: " + projectKey);
    } else {
      Logger.log("Could not extract project key from EPIC value: " + epicValue);
    }
  } else {
    Logger.log("No EPIC column found or no value in EPIC column");
  }
  
  // Set project key - prefer project derived from EPIC, fall back to tab settings
  ticketData.fields.project = {
    key: projectKey || tabSettings.jiraProject
  };
  
  // Log which project key is being used
  Logger.log("Using project key: " + ticketData.fields.project.key + 
             (projectKey ? " (derived from EPIC)" : " (from tab settings)"));
  
  // Check if project key is actually set
  if (!ticketData.fields.project.key) {
    var debugMsg = "WARNING: No project key found from EPIC or tab settings!";
    Logger.log(debugMsg);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      debugMsg,
      'Missing Project Key',
      15
    );
  }
  
  // Debug column mapping as we go
  var mappingDebug = [];
  
  // First pass: handle direct field mappings (including summary and description)
  // We need these fields to exist before we can modify them with prepend/append
  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    var value = rowData[i];
    
    if (!value) {
      mappingDebug.push('Column "' + header + '": Empty value, skipping');
      continue;
    }
    
    if (!columnMapping[header]) {
      mappingDebug.push('Column "' + header + '": No mapping defined, skipping');
      continue;
    }
    
    // Get the mapped field for this column (handle legacy format)
    var mappingConfig = columnMapping[header];
    var jiraField;
    
    if (typeof mappingConfig === 'object' && mappingConfig.mapped) {
      jiraField = mappingConfig.mapped;
      mappingDebug.push('Column "' + header + '": Mapped to "' + jiraField + '" (required=' + !!mappingConfig.required + ')');
    } else if (typeof mappingConfig === 'string') {
      // Legacy format - direct field mapping
      jiraField = mappingConfig;
      mappingDebug.push('Column "' + header + '": Legacy mapping to "' + jiraField + '"');
    } else {
      mappingDebug.push('Column "' + header + '": Invalid mapping format, skipping');
      continue;
    }
    
    // Skip prepend/append fields on first pass
    if (jiraField.startsWith('prepend_') || jiraField.startsWith('append_')) {
      mappingDebug.push('Column "' + header + '": Postponing prepend/append field for second pass');
      continue;
    }
    
    // Track columns mapped to ticketId
    if (jiraField === 'ticketId') {
      ticketIdColumns.push({
        header: header,
        index: i
      });
      mappingDebug.push('Column "' + header + '": Mapped to Ticket ID (will be updated after ticket creation)');
      continue; // Skip further processing of ticketId mapping (this is for output only)
    }
    
    // Handle different types of fields
    switch (jiraField) {
      case 'summary':
        ticketData.fields.summary = value;
        mappingDebug.push('Column "' + header + '": Set summary to "' + value + '"');
        break;
        
      case 'description':
        ticketData.fields.description = value;
        mappingDebug.push('Column "' + header + '": Set description to "' + value + '"');
        break;
        
      case 'issuetype':
        ticketData.fields.issuetype.name = value;
        mappingDebug.push('Column "' + header + '": Set issue type to "' + value + '"');
        break;
        
      case 'priority':
        ticketData.fields.priority = { name: value };
        mappingDebug.push('Column "' + header + '": Set priority to "' + value + '"');
        break;
        
      case 'labels':
        // Split comma-separated labels
        if (typeof value === 'string') {
          ticketData.fields.labels = value.split(',').map(function(label) {
            return label.trim();
          });
          mappingDebug.push('Column "' + header + '": Set labels to ' + JSON.stringify(ticketData.fields.labels));
        } else if (Array.isArray(value)) {
          ticketData.fields.labels = value;
          mappingDebug.push('Column "' + header + '": Set labels to array: ' + JSON.stringify(value));
        }
        break;
        
      case 'epicLink':
        // If EPIC field ID is configured and value exists, link to EPIC
        if (epicFieldId && value) {
          ticketData.fields[epicFieldId] = value;
          mappingDebug.push('Column "' + header + '": Set Epic Link to "' + value + '" using field ' + epicFieldId);
          
          // Try to extract project from this epic value as well
          if (!projectKey) {
            var epicParts = value.split('-');
            if (epicParts.length > 1) {
              projectKey = epicParts[0];
              // Update the project key
              ticketData.fields.project.key = projectKey;
              mappingDebug.push('Column "' + header + '": Updated project key to "' + projectKey + '" from mapped Epic Link');
            }
          }
        } else {
          if (!epicFieldId) {
            mappingDebug.push('Column "' + header + '": WARNING - Epic field ID not configured but column is mapped to epicLink');
          }
        }
        break;
        
      case 'components':
        // Handle multiple components as comma-separated values
        if (typeof value === 'string') {
          var components = value.split(',').map(function(component) {
            return { name: component.trim() };
          });
          ticketData.fields.components = components;
          mappingDebug.push('Column "' + header + '": Set components to ' + JSON.stringify(components));
        }
        break;
        
      case 'assignee':
        ticketData.fields.assignee = { name: value };
        mappingDebug.push('Column "' + header + '": Set assignee to "' + value + '"');
        break;
        
      case 'reporter':
        ticketData.fields.reporter = { name: value };
        mappingDebug.push('Column "' + header + '": Set reporter to "' + value + '"');
        break;
        
      case 'duedate':
        // Format date as YYYY-MM-DD if it's a Date object
        if (value instanceof Date) {
          ticketData.fields.duedate = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          mappingDebug.push('Column "' + header + '": Set due date to "' + ticketData.fields.duedate + '" (converted from Date object)');
        } else {
          ticketData.fields.duedate = value;
          mappingDebug.push('Column "' + header + '": Set due date to "' + value + '"');
        }
        break;
        
      default:
        mappingDebug.push('Column "' + header + '": Unhandled field type "' + jiraField + '", skipping');
        break;
    }
  }
  
  // Second pass: handle prepend/append fields
  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    var value = rowData[i];
    
    if (!value) continue;
    if (!columnMapping[header]) continue;
    
    // Get the mapped field for this column
    var mappingConfig = columnMapping[header];
    var jiraField;
    var separator = ' '; // Default separator
    
    if (typeof mappingConfig === 'object' && mappingConfig.mapped) {
      jiraField = mappingConfig.mapped;
      
      // Get custom separator if specified
      if (mappingConfig.separator) {
        separator = mappingConfig.separator;
      }
    } else if (typeof mappingConfig === 'string') {
      // Legacy format - direct field mapping
      jiraField = mappingConfig;
    } else {
      continue;
    }
    
    // Only process prepend/append fields in second pass
    if (!jiraField.startsWith('prepend_') && !jiraField.startsWith('append_')) {
      continue;
    }
    
    // Handle prepend/append to summary
    if (jiraField === 'prepend_summary') {
      // Initialize summary if it doesn't exist
      if (!ticketData.fields.summary) {
        ticketData.fields.summary = value;
        mappingDebug.push('Column "' + header + '": Set summary to "' + value + '" (no previous value)');
      } else {
        ticketData.fields.summary = value + separator + ticketData.fields.summary;
        mappingDebug.push('Column "' + header + '": Prepended "' + value + '" to summary with separator "' + separator + '"');
      }
    }
    else if (jiraField === 'append_summary') {
      // Initialize summary if it doesn't exist
      if (!ticketData.fields.summary) {
        ticketData.fields.summary = value;
        mappingDebug.push('Column "' + header + '": Set summary to "' + value + '" (no previous value)');
      } else {
        ticketData.fields.summary = ticketData.fields.summary + separator + value;
        mappingDebug.push('Column "' + header + '": Appended "' + value + '" to summary with separator "' + separator + '"');
      }
    }
    // Handle prepend/append to description
    else if (jiraField === 'prepend_description') {
      // Initialize description if it doesn't exist
      if (!ticketData.fields.description) {
        ticketData.fields.description = value;
        mappingDebug.push('Column "' + header + '": Set description to "' + value + '" (no previous value)');
      } else {
        ticketData.fields.description = value + separator + ticketData.fields.description;
        mappingDebug.push('Column "' + header + '": Prepended "' + value + '" to description with separator "' + separator + '"');
      }
    }
    else if (jiraField === 'append_description') {
      // Initialize description if it doesn't exist
      if (!ticketData.fields.description) {
        ticketData.fields.description = value;
        mappingDebug.push('Column "' + header + '": Set description to "' + value + '" (no previous value)');
      } else {
        ticketData.fields.description = ticketData.fields.description + separator + value;
        mappingDebug.push('Column "' + header + '": Appended "' + value + '" to description with separator "' + separator + '"');
      }
    }
  }
  
  // Store ticketIdColumns in the ticketData for use after ticket creation
  if (ticketIdColumns.length > 0) {
    // We can't include this in fields as Jira API won't accept it, so store it at top level
    ticketData._ticketIdColumns = ticketIdColumns;
    mappingDebug.push('Stored ' + ticketIdColumns.length + ' ticketId columns for updating after ticket creation');
  }
  
  // Log all field mappings for debugging
  Logger.log('Field mapping details:');
  mappingDebug.forEach(function(msg) {
    Logger.log('  ' + msg);
  });
  
  // Check if required fields are present
  if (!ticketData.fields.summary) {
    var errorMsg = 'Summary field is required';
    Logger.log('ERROR: ' + errorMsg);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      errorMsg,
      'Missing Required Field',
      10
    );
    throw new Error(errorMsg);
  }
  
  // Final check to ensure we have a project key
  if (!ticketData.fields.project.key) {
    var errorMsg = 'No project key found from EPIC or tab settings';
    Logger.log('ERROR: ' + errorMsg);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      errorMsg,
      'Missing Project Key',
      10
    );
    throw new Error(errorMsg);
  }
  
  // Log the final ticket data payload
  Logger.log('Final ticket data: ' + JSON.stringify(ticketData));
  
  return ticketData;
}

/**
 * Create a Jira ticket with the provided data
 */
function createJiraTicket(ticketData) {
  var response = makeJiraRequest('issue', 'POST', ticketData);
  
  if (response.getResponseCode() !== 201) {
    throw new Error('Failed to create ticket: ' + response.getContentText());
  }
  
  return response;
}

/**
 * Sync Jira ticket status back to the sheet
 */
function syncJiraStatus() {
  if (!isJiraConfigured()) {
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert(
      'Jira Settings Required',
      'Your Jira integration is not configured yet. Would you like to configure it now?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      showSettingsDialog();
    }
    return;
  }
  
  // Show toast that sync is starting
  showToast("Starting Jira ticket status sync...", "Syncing", 3);
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var ticketKeyColumn = findTicketKeyColumn(headers);
  var statusColumn = findColumnByMapping(headers, 'status', 'Status');
  var numberColumn = headers.indexOf('#') + 1;
  
  if (ticketKeyColumn < 1 || statusColumn < 1) {
    SpreadsheetApp.getUi().alert('Sheet must contain "Ticket Key" and "Status" columns!');
    showToast("Sync failed: Missing required columns", "Error", 5);
    return;
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 2) {
    showToast("No data rows found to sync", "Warning", 5);
    return; // Only header rows exist
  }
  
  var ticketKeys = sheet.getRange(3, ticketKeyColumn, lastRow - 2, 1).getValues();
  var updatedCount = 0;
  var errorCount = 0;
  
  for (var i = 0; i < ticketKeys.length; i++) {
    var ticketKey = ticketKeys[i][0];
    if (!ticketKey) continue;
    
    try {
      // Show progress toast every 5 tickets
      if (i % 5 === 0) {
        showToast("Syncing ticket " + (i+1) + " of " + ticketKeys.length, "Sync Progress", 2);
      }
      
      var response = makeJiraRequest('issue/' + ticketKey, 'GET');
      if (response.getResponseCode() === 200) {
        var issueData = JSON.parse(response.getContentText());
        var status = issueData.fields.status.name;
        
        sheet.getRange(i + 3, statusColumn).setValue(status);
        
        // Update or add the hyperlink in the # column
        if (numberColumn > 0) {
          updateCellWithTicketLink(sheet, i + 3, numberColumn, ticketKey);
        }
        
        updatedCount++;
      } else {
        errorCount++;
        showToast("Error syncing " + ticketKey + ": Status " + response.getResponseCode(), "Error", 5);
      }
    } catch (error) {
      errorCount++;
      showToast("Error syncing " + ticketKey + ": " + error.toString().substring(0, 50), "Error", 5);
    }
    
    // Flush changes every 10 tickets to prevent timeout
    if (i % 10 === 0) {
      SpreadsheetApp.flush();
    }
  }
  
  // Final toast with summary
  if (updatedCount > 0) {
    showToast(updatedCount + " ticket(s) updated successfully" + (errorCount > 0 ? ", " + errorCount + " errors" : ""), "Sync Complete", 10);
  } else if (errorCount > 0) {
    showToast("Sync failed with " + errorCount + " errors", "Sync Failed", 10);
  }
  
  SpreadsheetApp.getUi().alert(updatedCount + ' ticket status(es) updated.\n' + 
                              errorCount + ' error(s) occurred.');
}

/**
 * Installs the onEdit trigger for detecting row insertions
 *
 * @since TBD
 */
function installEditTrigger() {
  // Check if trigger already exists
  var triggers = ScriptApp.getProjectTriggers();
  var triggerExists = false;
  
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onSheetEdit') {
      triggerExists = true;
      break;
    }
  }
  
  if (!triggerExists) {
    // Create trigger for onSheetEdit function
    ScriptApp.newTrigger('onSheetEdit')
      .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
      .onEdit()
      .create();
    
    SpreadsheetApp.getUi().alert('Edit trigger installed successfully!');
  } else {
    SpreadsheetApp.getUi().alert('Edit trigger is already installed.');
  }
}

/**
 * Checks if a sheet has the required columns for Jira sync
 *
 * @since TBD
 * @param {Sheet} sheet The sheet to check
 * @return {boolean} True if all required columns are present, false otherwise
 */
function hasRequiredColumns(sheet) {
  try {
    // Get headers from row 2 (not row 1)
    var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Log all headers for debugging
    Logger.log('Headers found: ' + JSON.stringify(headers));
    
    // Get column mapping to check for ticketId mapping
    var columnMapping = getColumnMapping();
    
    // Check for essential columns: Ticket Key and Status
    var ticketKeyIdx = findTicketKeyColumn(headers, columnMapping) - 1; // Convert back to 0-based index
    var statusIdx = findColumnByMapping(headers, 'status', 'Status', columnMapping) - 1; // Convert back to 0-based index
    var hasTicketKey = ticketKeyIdx >= 0;
    var hasStatus = statusIdx >= 0;
    
    // Log essential column checks
    logToSheet('hasRequiredColumns', 'info', 'Essential column check', {
      ticketKeyIdx: ticketKeyIdx,
      hasTicketKey: hasTicketKey,
      statusIdx: statusIdx,
      hasStatus: hasStatus
    });
    
    // Check for EPIC column (important for project organization)
    var epicIdx = findColumnByMapping(headers, 'epicLink', ['EPIC', 'Epic'], columnMapping);
    var hasEpicColumn = epicIdx > 0;
    
    logToSheet('hasRequiredColumns', 'info', 'Epic column check', {
      epicIdx: epicIdx,
      hasEpicColumn: hasEpicColumn
    });
    
    // Check for required mapped columns using the column mapping's required flag
    var hasAllRequiredMappedColumns = true;
    var hasMappedColumn = false;
    var missingRequiredColumns = [];
    
    // Log column mapping
    logToSheet('hasRequiredColumns', 'info', 'Column mapping', columnMapping);
    
    // Check if we have any mapped columns and all required ones are present
    for (var field in columnMapping) {
      if (columnMapping[field] && typeof columnMapping[field] === 'object' && columnMapping[field].mapped) {
        hasMappedColumn = true;
        
        // If this is a required column, check if it exists in the headers
        if (columnMapping[field].required && headers.indexOf(field) < 0) {
          hasAllRequiredMappedColumns = false;
          missingRequiredColumns.push(field);
        }
      }
    }
    
    logToSheet('hasRequiredColumns', 'info', 'Required mapping check', {
      hasMappedColumn: hasMappedColumn,
      hasAllRequiredMappedColumns: hasAllRequiredMappedColumns,
      missingRequiredColumns: missingRequiredColumns
    });
    
    // Get tab settings
    var tabSettings = getTabSettings(sheet.getName());
    var hasProjectKey = tabSettings.jiraProject && tabSettings.jiraProject.trim() !== '';
    var hasEpicFieldId = tabSettings.epicFieldId && tabSettings.epicFieldId.trim() !== '';
    
    // Log tab settings checks
    logToSheet('hasRequiredColumns', 'info', 'Tab settings check', {
      tabName: sheet.getName(),
      projectKey: tabSettings.jiraProject || 'NOT SET',
      epicFieldId: tabSettings.epicFieldId || 'NOT SET'
    });
    
    // Create a detailed debug message and log it
    var debugResult = 'Column check results:\n' + 
                      '- Has Ticket Key column or ticketId mapping: ' + hasTicketKey + '\n' +
                      '- Has Status column: ' + hasStatus + '\n' +
                      '- Has Epic column: ' + hasEpicColumn + '\n' +
                      '- Has mapped columns: ' + hasMappedColumn + '\n' +
                      '- Has all required mapped columns: ' + hasAllRequiredMappedColumns + 
                      (missingRequiredColumns.length > 0 ? ' (Missing: ' + missingRequiredColumns.join(', ') + ')' : '') + '\n' +
                      '- Has project key in tab settings: ' + hasProjectKey + '\n' +
                      '- Has Epic Field ID in tab settings: ' + hasEpicFieldId;
    
    logToSheet('hasRequiredColumns', 'info', 'Column check summary', debugResult);
    
    // Display debug toast for 30 seconds
    SpreadsheetApp.getActiveSpreadsheet().toast(
      debugResult,
      'Column Check Debug',
      30
    );
    
    // Final result - allow using EPIC column as an alternative to project key in settings
    var hasProjectSource = hasProjectKey || hasEpicColumn;
    
    var result = hasTicketKey && hasStatus && hasMappedColumn && hasAllRequiredMappedColumns && 
                hasProjectSource && (hasEpicFieldId || !hasEpicColumn); // Need Epic Field ID if Epic Column exists
    
    logToSheet('hasRequiredColumns', 'info', 'Final result', {
      result: result,
      hasTicketKey: hasTicketKey,
      hasStatus: hasStatus,
      hasMappedColumn: hasMappedColumn,
      hasAllRequiredMappedColumns: hasAllRequiredMappedColumns,
      hasProjectKey: hasProjectKey,
      hasEpicColumn: hasEpicColumn,
      hasProjectSource: hasProjectSource,
      hasEpicFieldId: hasEpicFieldId
    });
    
    return result;
  } catch (error) {
    Logger.log("Error checking required columns: " + error);
    
    // Show error in toast
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Error checking required columns: " + error,
      'Column Check Error',
      15
    );
    
    // Log to debug sheet
    logToSheet('hasRequiredColumns', 'error', 'Column check error', {
      error: error.toString(),
      stack: error.stack
    });
    
    return false;
  }
}

/**
 * Handler for edit events, specifically looking for row insertions
 * in sheets other than the first one
 *
 * @since TBD
 * @param {Object} e The event object from the trigger
 */
function onSheetEdit(e) {
  // Get the active sheet
  var sheet = e.source.getActiveSheet();
  var sheets = e.source.getSheets();
  
  // Log to debug sheet with edit event details
  logToSheet('onSheetEdit', 'info', 'Edit detected', {
    sheet: sheet.getName(),
    row: e.range.getRow(),
    column: e.range.getColumn(),
    value: e.value,
    numRows: e.range.getNumRows(),
    numCols: e.range.getNumColumns()
  });
  
  // Show detailed debug toast
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Edit detected: Sheet=' + sheet.getName() + 
    ', Row=' + e.range.getRow() + 
    ', Col=' + e.range.getColumn() + 
    ', Value=' + e.value,
    'Debug: Edit Event',
    15
  );
  
  // Check if this is not the first sheet
  if (sheet.getIndex() !== 1) {
    // Check for specific actions that indicate row insertion or modification
    if (e.range && e.range.getRow() > 2) { // Start after header rows
      // Check if Jira is configured
      var isConfigured = isJiraConfigured();
      logToSheet('onSheetEdit', 'info', 'Jira configuration check', {
        configured: isConfigured
      });
      
      if (!isConfigured) {
        // Log configuration issue
        logToSheet('onSheetEdit', 'warning', 'Jira not configured', 'User needs to configure Jira settings');
        
        // Wait a moment for the UI to stabilize
        SpreadsheetApp.getActiveSpreadsheet().toast(
          'Edit detected. Please configure Jira settings to enable ticket creation.',
          'Jira Setup Required',
          -1
        );
        
        // Show configuration dialog
        var ui = SpreadsheetApp.getUi();
        var response = ui.alert(
          'Jira Settings Required',
          'You\'re editing a row, but your Jira integration is not configured yet. Would you like to configure it now?',
          ui.ButtonSet.YES_NO
        );
        
        if (response === ui.Button.YES) {
          showSettingsDialog();
        }
      } else {
        // Check required columns and log the result
        var hasRequiredCols = hasRequiredColumns(sheet);
        logToSheet('onSheetEdit', 'info', 'Required columns check', {
          sheet: sheet.getName(),
          hasAllRequired: hasRequiredCols
        });
        
        if (hasRequiredCols) {
          // Check if this row already has a ticket
          var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
          var ticketKeyColumn = findTicketKeyColumn(headers);
          var row = e.range.getRow();
          
          logToSheet('onSheetEdit', 'info', 'Checking row for ticket', {
            row: row,
            ticketKeyColumn: ticketKeyColumn
          });
          
          // If we found the Ticket Key column, check if it has a value
          if (ticketKeyColumn > 0) {
            var ticketKey = sheet.getRange(row, ticketKeyColumn).getValue();
            
            logToSheet('onSheetEdit', 'info', 'Ticket key check', {
              row: row,
              hasTicket: !!ticketKey,
              ticketKey: ticketKey || 'none'
            });
            
            // If no ticket key yet, validate the row for required fields
            if (!ticketKey) {
              // Validate the row
              var validation = validateRowForTicket(sheet, row);
              
              // Log validation results in detail
              logToSheet('onSheetEdit', 'info', 'Row validation result', {
                row: row,
                success: validation.success,
                errors: validation.errors,
                missingFields: validation.missingRequiredFields
              });
              
              if (validation.success) {
                // All required fields are present - offer to create ticket
                var ui = SpreadsheetApp.getUi();
                var response = ui.alert(
                  'Create Jira Ticket?',
                  'This row has all required fields but no Jira ticket yet. Create one now?',
                  ui.ButtonSet.YES_NO
                );
                
                logToSheet('onSheetEdit', 'info', 'Create ticket prompt', {
                  row: row,
                  userResponse: response === ui.Button.YES ? 'YES' : 'NO'
                });
                
                if (response === ui.Button.YES) {
                  // Select just this row and create a ticket
                  sheet.getRange(row, 1, 1, sheet.getLastColumn()).activate();
                  createJiraTicketsFromSelection();
                }
              } else {
                // Log validation failure with details
                logToSheet('onSheetEdit', 'error', 'Row has missing required fields', {
                  row: row,
                  errors: validation.errors,
                  missingFields: validation.missingRequiredFields
                });
                
                // Missing required fields - show alert with details
                var ui = SpreadsheetApp.getUi();
                var message = 'Cannot create Jira ticket due to the following issues:\n\n' + 
                              validation.errors.join('\n\n');
                
                // Show a detailed alert
                ui.alert('Missing Required Fields', message, ui.ButtonSet.OK);
              }
            }
          } else {
            logToSheet('onSheetEdit', 'warning', 'No Ticket Key column found', {
              sheet: sheet.getName(),
              headers: headers
            });
            
            // Show more helpful message to user
            SpreadsheetApp.getActiveSpreadsheet().toast(
              'This sheet needs a "Ticket Key" column or a column mapped to "Jira Ticket ID" to store created tickets. ' +
              'Use the "Configure Column Mapping" option from the Jira Integration menu.',
              'Column Setup Needed',
              15
            );
          }
        } else {
          logToSheet('onSheetEdit', 'warning', 'Sheet missing required columns', {
            sheet: sheet.getName()
          });
        }
      }
    }
  }
}

/**
 * Get the current active sheet name
 * 
 * @since TBD
 * @return {string} Name of the active sheet.
 */
function getActiveSheetName() {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    if (sheet) {
      return sheet.getName();
    }
  } catch (error) {
    Logger.log("Error getting active sheet name: " + error);
  }
  
  return "Sheet1"; // Default fallback
}

/**
 * Get the Epic field ID for the active tab
 * This handles the customfield ID that Jira uses for Epic Links
 *
 * @since TBD
 * @return {string} The Epic field ID
 */
function getEpicFieldId() {
  var tabSettings = getTabSettings();
  return tabSettings.epicFieldId || 'customfield_10014'; // Default ID if not set
}

/**
 * Validates a row against required fields for ticket creation
 * 
 * @since TBD
 * @param {Sheet} sheet The sheet containing the row
 * @param {number} rowIndex The row index (1-based)
 * @return {Object} Validation result containing success flag and errors
 */
function validateRowForTicket(sheet, rowIndex) {
  var result = {
    success: true,
    errors: [],
    missingRequiredFields: []
  };
  
  try {
    // Log start of validation
    logToSheet('validateRowForTicket', 'info', 'Starting validation', {
      sheet: sheet.getName(),
      row: rowIndex
    });
    
    // Get headers and row data
    var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    var rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Log data for debugging
    logToSheet('validateRowForTicket', 'info', 'Retrieved row data', {
      headers: headers,
      rowData: rowData
    });
    
    // Check if row is empty
    var hasData = false;
    for (var i = 0; i < rowData.length; i++) {
      if (rowData[i]) {
        hasData = true;
        break;
      }
    }
    
    if (!hasData) {
      result.success = false;
      result.errors.push('Row is empty');
      logToSheet('validateRowForTicket', 'error', 'Row is empty', { row: rowIndex });
      return result;
    }
    
    // Get column mapping
    var columnMapping = getColumnMapping();
    var tabSettings = getTabSettings();
    
    // Log mapping and settings for debugging
    logToSheet('validateRowForTicket', 'info', 'Retrieved settings', {
      sheet: sheet.getName(),
      tabSettings: tabSettings,
      columnMapping: columnMapping
    });
    
    // Check project key (needed in tab settings or from EPIC)
    var projectKey = null;
    
    // Try to find Epic column for project extraction
    var epicColumnIndex = findColumnByMapping(headers, 'epicLink', ['EPIC', 'Epic'], columnMapping) - 1; // Convert back to 0-based
    
    logToSheet('validateRowForTicket', 'info', 'Epic column check', {
      epicColumnIndex: epicColumnIndex,
      hasEpicColumn: epicColumnIndex >= 0
    });
    
    if (epicColumnIndex >= 0 && rowData[epicColumnIndex]) {
      var epicValue = rowData[epicColumnIndex];
      var epicParts = epicValue.split('-');
      if (epicParts.length > 1) {
        projectKey = epicParts[0];
        logToSheet('validateRowForTicket', 'info', 'Project key extracted from Epic', {
          epic: epicValue,
          extractedProjectKey: projectKey
        });
      } else {
        logToSheet('validateRowForTicket', 'warning', 'Epic value does not contain project key', {
          epic: epicValue,
          expected: 'PROJECT-123 format'
        });
      }
    } else if (epicColumnIndex >= 0) {
      logToSheet('validateRowForTicket', 'warning', 'Epic column found but no value', {
        epicColumnIndex: epicColumnIndex
      });
    }
    
    // Check if we have a project key from somewhere
    if (!projectKey && (!tabSettings.jiraProject || tabSettings.jiraProject.trim() === '')) {
      result.success = false;
      result.errors.push('No project key available - check Epic value or tab settings');
      logToSheet('validateRowForTicket', 'error', 'No project key available', {
        epicValue: epicColumnIndex >= 0 ? rowData[epicColumnIndex] : 'No Epic column',
        tabProjectKey: tabSettings.jiraProject || 'Not set'
      });
    } else {
      // Use tab settings if no project key from Epic
      if (!projectKey) {
        projectKey = tabSettings.jiraProject;
        logToSheet('validateRowForTicket', 'info', 'Using project key from tab settings', {
          projectKey: projectKey
        });
      }
    }
    
    // Log column mapping checks
    logToSheet('validateRowForTicket', 'info', 'Checking required mapped fields', {
      mappedFields: Object.keys(columnMapping)
    });
    
    // Check required mapped fields
    for (var field in columnMapping) {
      if (columnMapping[field] && 
          typeof columnMapping[field] === 'object' && 
          columnMapping[field].mapped && 
          columnMapping[field].required) {
        
        var columnIndex = headers.indexOf(field);
        var isMissing = false;
        
        if (columnIndex < 0) {
          // Required field's column doesn't exist in the sheet
          isMissing = true;
          logToSheet('validateRowForTicket', 'error', 'Required column not found in sheet', {
            field: field,
            mappedTo: columnMapping[field].mapped
          });
        } else if (!rowData[columnIndex]) {
          // Column exists but value is empty
          isMissing = true;
          logToSheet('validateRowForTicket', 'error', 'Required field has no value', {
            field: field,
            columnIndex: columnIndex,
            mappedTo: columnMapping[field].mapped
          });
        } else {
          // Field has a value, log it
          logToSheet('validateRowForTicket', 'info', 'Required field check passed', {
            field: field,
            value: rowData[columnIndex],
            mappedTo: columnMapping[field].mapped
          });
        }
        
        if (isMissing) {
          result.success = false;
          result.missingRequiredFields.push({
            field: field,
            mappedTo: columnMapping[field].mapped
          });
        }
      }
    }
    
    // Always check Summary - it's required for all tickets
    logToSheet('validateRowForTicket', 'info', 'Checking for summary field');
    var summaryFound = false;
    
    // Check for directly mapped summary
    for (var field in columnMapping) {
      if (columnMapping[field] && 
          typeof columnMapping[field] === 'object' && 
          columnMapping[field].mapped === 'summary') {
        
        var columnIndex = headers.indexOf(field);
        if (columnIndex >= 0 && rowData[columnIndex]) {
          summaryFound = true;
          logToSheet('validateRowForTicket', 'info', 'Summary found in mapped field', {
            field: field,
            value: rowData[columnIndex]
          });
          break;
        }
      }
    }
    
    // Check for old-style "Summary" column
    if (!summaryFound) {
      var summaryIndex = headers.indexOf('Summary');
      if (summaryIndex >= 0 && rowData[summaryIndex]) {
        summaryFound = true;
        logToSheet('validateRowForTicket', 'info', 'Summary found in standard column', {
          value: rowData[summaryIndex]
        });
      }
    }
    
    if (!summaryFound) {
      result.success = false;
      result.missingRequiredFields.push({
        field: 'Summary',
        mappedTo: 'summary'
      });
      logToSheet('validateRowForTicket', 'error', 'No summary field found', {
        mappedFields: Object.keys(columnMapping)
      });
    }
    
    // Generate error messages for missing fields
    if (result.missingRequiredFields.length > 0) {
      var fieldList = result.missingRequiredFields.map(function(f) {
        return f.field + ' (→ ' + f.mappedTo + ')';
      }).join(', ');
      
      result.errors.push('Missing required fields: ' + fieldList);
      logToSheet('validateRowForTicket', 'error', 'Missing required fields', {
        fields: result.missingRequiredFields
      });
    }
    
    // Log validation result
    logToSheet('validateRowForTicket', 'info', 'Validation completed', {
      success: result.success,
      errors: result.errors,
      missingFields: result.missingRequiredFields
    });
    
    return result;
  
  } catch (error) {
    logToSheet('validateRowForTicket', 'error', 'Validation error', {
      error: error.toString(),
      stack: error.stack
    });
    
    result.success = false;
    result.errors.push('Validation error: ' + error.toString());
    return result;
  }
}

/**
 * Creates or gets the debug sheet for logging
 *
 * @since TBD
 * @return {Sheet} The debug sheet
 */
function getDebugSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var debugSheet;
  
  try {
    debugSheet = ss.getSheetByName('JiraDebug');
    
    if (!debugSheet) {
      // Create a new debug sheet
      debugSheet = ss.insertSheet('JiraDebug');
      
      // Setup headers
      debugSheet.getRange('A1:E1').setValues([['Timestamp', 'Source', 'Type', 'Message', 'Details']]);
      debugSheet.getRange('A1:E1').setFontWeight('bold');
      debugSheet.setFrozenRows(1);
      
      // Set column widths
      debugSheet.setColumnWidth(1, 150); // Timestamp
      debugSheet.setColumnWidth(2, 100); // Source
      debugSheet.setColumnWidth(3, 80);  // Type
      debugSheet.setColumnWidth(4, 300); // Message
      debugSheet.setColumnWidth(5, 500); // Details
    }
    
    return debugSheet;
  } catch (error) {
    // If there's an error, try to log it in toast
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Error creating debug sheet: ' + error.toString(),
      'Debug Error',
      10
    );
    
    return null;
  }
}

/**
 * Logs a message to the debug sheet
 *
 * @since TBD
 * @param {string} source The source of the message (function name)
 * @param {string} type The type of message (info, warning, error)
 * @param {string} message The message to log
 * @param {Object|string} details Optional details (will be converted to string)
 */
function logToSheet(source, type, message, details) {
  var debugSheet = getDebugSheet();
  
  if (!debugSheet) {
    return; // Can't log if debug sheet not available
  }
  
  try {
    // Format details
    var detailsStr = '';
    if (details) {
      if (typeof details === 'object') {
        detailsStr = JSON.stringify(details, null, 2);
      } else {
        detailsStr = details.toString();
      }
    }
    
    // Format timestamp
    var now = new Date();
    var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    
    // Add row to debug sheet
    debugSheet.appendRow([timestamp, source, type, message, detailsStr]);
    
    // Highlight errors
    var lastRow = debugSheet.getLastRow();
    if (type.toLowerCase() === 'error') {
      debugSheet.getRange(lastRow, 1, 1, 5).setBackground('#ffcccc');
    } else if (type.toLowerCase() === 'warning') {
      debugSheet.getRange(lastRow, 1, 1, 5).setBackground('#ffffcc');
    } else if (type.toLowerCase() === 'success') {
      debugSheet.getRange(lastRow, 1, 1, 5).setBackground('#ccffcc');
    }
    
    // Auto-resize columns if needed
    if (lastRow % 50 === 0) {
      debugSheet.autoResizeColumns(1, 5);
    }
  } catch (error) {
    // If failed to log to sheet, try toast as fallback
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Failed to log to debug sheet: ' + error.toString(),
      'Log Error',
      10
    );
  }
}

/**
 * Clears the debug sheet except for the header row
 * 
 * @since TBD
 */
function clearDebugSheet() {
  var debugSheet = getDebugSheet();
  
  if (!debugSheet) {
    return;
  }
  
  var lastRow = debugSheet.getLastRow();
  
  if (lastRow > 1) {
    debugSheet.deleteRows(2, lastRow - 1);
  }
  
  // Log that the sheet was cleared
  logToSheet('clearDebugSheet', 'info', 'Debug sheet has been cleared', '');
}

/**
 * Shows the debug sheet and brings it to front
 * 
 * @since TBD
 */
function showDebugSheet() {
  var debugSheet = getDebugSheet();
  if (debugSheet) {
    debugSheet.activate();
    logToSheet('showDebugSheet', 'info', 'Debug sheet activated', '');
  }
}

/**
 * Shows a dialog with the current column mapping
 * 
 * @since TBD
 */
function viewColumnMapping() {
  var mapping = getColumnMapping();
  
  // Format the mapping for display
  var formattedMapping = [];
  
  for (var column in mapping) {
    if (mapping[column]) {
      var mapObj = mapping[column];
      var mappedField = '';
      var isRequired = false;
      var separator = '';
      
      if (typeof mapObj === 'object') {
        mappedField = mapObj.mapped || 'Unknown';
        isRequired = mapObj.required || false;
        separator = mapObj.separator || '';
      } else {
        // Handle legacy string mapping
        mappedField = mapObj;
      }
      
      formattedMapping.push({
        column: column,
        mappedTo: mappedField,
        required: isRequired ? 'Yes' : 'No',
        separator: separator
      });
    }
  }
  
  // Log to debug sheet
  logToSheet('viewColumnMapping', 'info', 'Column mapping viewed', mapping);
  
  // Sort by column name
  formattedMapping.sort(function(a, b) {
    return a.column.localeCompare(b.column);
  });
  
  // Create HTML output
  var html = '<html><body style="font-family: Arial, sans-serif; padding: 20px;">';
  html += '<h2>Column Mapping Configuration</h2>';
  
  if (formattedMapping.length === 0) {
    html += '<p>No column mapping configured yet.</p>';
  } else {
    html += '<table border="1" cellpadding="4" cellspacing="0" style="border-collapse: collapse;">';
    html += '<tr style="background-color: #f2f2f2;"><th>Sheet Column</th><th>Jira Field</th><th>Required</th><th>Separator</th></tr>';
    
    for (var i = 0; i < formattedMapping.length; i++) {
      var item = formattedMapping[i];
      var rowStyle = i % 2 === 0 ? 'background-color: #f9f9f9;' : '';
      var requiredStyle = item.required === 'Yes' ? 'color: red; font-weight: bold;' : '';
      
      html += '<tr style="' + rowStyle + '">';
      html += '<td>' + item.column + '</td>';
      html += '<td>' + item.mappedTo + '</td>';
      html += '<td style="' + requiredStyle + '">' + item.required + '</td>';
      html += '<td>' + (item.separator ? '"' + item.separator + '"' : '') + '</td>';
      html += '</tr>';
    }
    
    html += '</table>';
  }
  
  html += '<p><i>Note: Column mapping is stored globally across all sheets in this spreadsheet.</i></p>';
  html += '</body></html>';
  
  // Show the dialog
  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(400)
    .setTitle('Column Mapping Details');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Column Mapping Details');
}

/**
 * Tests row validation on the current row
 * 
 * @since TBD
 */
function testRowValidation() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var activeCell = sheet.getActiveCell();
  var row = activeCell.getRow();
  
  // Skip header rows
  if (row <= 2) {
    SpreadsheetApp.getUi().alert('Please select a data row (row 3 or later)');
    return;
  }
  
  // Validate the row
  logToSheet('testRowValidation', 'info', 'Testing row validation', {
    sheet: sheet.getName(),
    row: row
  });
  
  var validation = validateRowForTicket(sheet, row);
  
  // Create HTML output for validation results
  var html = '<html><body style="font-family: Arial, sans-serif; padding: 20px;">';
  html += '<h2>Row Validation Results</h2>';
  html += '<p><b>Sheet:</b> ' + sheet.getName() + '</p>';
  html += '<p><b>Row:</b> ' + row + '</p>';
  html += '<p><b>Valid for ticket creation:</b> <span style="color: ' + (validation.success ? 'green' : 'red') + ';">' + 
          (validation.success ? 'YES' : 'NO') + '</span></p>';
  
  if (!validation.success) {
    html += '<h3 style="color: red;">Errors:</h3>';
    html += '<ul>';
    for (var i = 0; i < validation.errors.length; i++) {
      html += '<li>' + validation.errors[i] + '</li>';
    }
    html += '</ul>';
    
    if (validation.missingRequiredFields.length > 0) {
      html += '<h3 style="color: red;">Missing Required Fields:</h3>';
      html += '<table border="1" cellpadding="4" cellspacing="0" style="border-collapse: collapse;">';
      html += '<tr style="background-color: #f2f2f2;"><th>Sheet Column</th><th>Jira Field</th></tr>';
      
      for (var i = 0; i < validation.missingRequiredFields.length; i++) {
        var field = validation.missingRequiredFields[i];
        html += '<tr>';
        html += '<td>' + field.field + '</td>';
        html += '<td>' + field.mappedTo + '</td>';
        html += '</tr>';
      }
      
      html += '</table>';
    }
  }
  
  html += '<p><i>Full details available in the Debug sheet.</i></p>';
  html += '</body></html>';
  
  // Show the dialog
  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(400)
    .setTitle('Row Validation Results');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Row Validation Results');
}

/**
 * Shows the tab settings for the current tab
 * 
 * @since TBD
 */
function viewTabSettings() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var tabName = sheet.getName();
  var tabSettings = getTabSettings(tabName);
  
  // Check if the sheet has an EPIC column
  var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  var columnMapping = getColumnMapping();
  var epicColumn = findColumnByMapping(headers, 'epicLink', ['EPIC', 'Epic'], columnMapping);
  var hasEpicColumn = epicColumn > 0;
  
  // Log to debug sheet
  logToSheet('viewTabSettings', 'info', 'Tab settings viewed', {
    tab: tabName,
    settings: tabSettings,
    hasEpicColumn: hasEpicColumn,
    epicColumn: epicColumn
  });
  
  // Create HTML output
  var html = '<html><body style="font-family: Arial, sans-serif; padding: 20px;">';
  html += '<h2>Tab Settings: ' + tabName + '</h2>';
  
  html += '<table border="1" cellpadding="4" cellspacing="0" style="border-collapse: collapse; margin-bottom: 20px;">';
  html += '<tr style="background-color: #f2f2f2;"><th>Setting</th><th>Value</th></tr>';
  
  // Project key
  var projectKeyStyle = tabSettings.jiraProject ? '' : 'color: red;';
  html += '<tr>';
  html += '<td><b>Jira Project Key</b></td>';
  html += '<td style="' + projectKeyStyle + '">' + (tabSettings.jiraProject || 'NOT SET') + '</td>';
  html += '</tr>';
  
  // Default issue type
  html += '<tr>';
  html += '<td><b>Default Issue Type</b></td>';
  html += '<td>' + (tabSettings.defaultIssueType || 'Task') + '</td>';
  html += '</tr>';
  
  // Epic field ID
  var epicFieldStyle = tabSettings.epicFieldId ? '' : 'color: red;';
  html += '<tr>';
  html += '<td><b>Epic Field ID</b></td>';
  html += '<td style="' + epicFieldStyle + '">' + (tabSettings.epicFieldId || 'NOT SET') + '</td>';
  html += '</tr>';
  
  html += '</table>';
  
  html += '<p><b>Global Jira URL:</b> ' + getSettings().jiraUrl + '</p>';
  
  if (!tabSettings.jiraProject) {
    if (hasEpicColumn) {
      html += '<p style="color: green;"><b>Note:</b> No project key is set for this tab, but an EPIC column was found. ' +
              'Project key will be determined from EPIC values (e.g., "PROJECT-123").</p>';
    } else {
      html += '<p style="color: red;"><b>Warning:</b> No project key is set for this tab. You must either:</p>';
      html += '<ul>';
      html += '<li>Configure a project key in tab settings, or</li>';
      html += '<li>Have an EPIC column with values in the format "PROJECT-123"</li>';
      html += '</ul>';
    }
  }
  
  if (!tabSettings.epicFieldId) {
    html += '<p style="color: red;"><b>Warning:</b> No Epic Field ID is set. This is needed if you want to link tickets to Epics.</p>';
  }
  
  html += '<p><i>Update these settings using the "Configure Settings" option in the Jira Integration menu.</i></p>';
  html += '</body></html>';
  
  // Show the dialog
  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(400)
    .setTitle('Tab Settings');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Tab Settings');
}

/**
 * Analyzes the sheet structure and reports missing columns
 * 
 * @since TBD
 */
function analyzeSheetColumns() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  var columnMapping = getColumnMapping();
  var tabSettings = getTabSettings();
  
  // Check if Epic column exists in the sheet
  var epicColumn = findColumnByMapping(headers, 'epicLink', ['EPIC', 'Epic'], columnMapping);
  var hasEpicColumn = epicColumn > 0;
  
  // Log analysis start
  logToSheet('analyzeSheetColumns', 'info', 'Analyzing sheet columns', {
    sheet: sheet.getName(),
    headers: headers,
    columnMapping: columnMapping,
    tabSettings: tabSettings,
    hasEpicColumn: hasEpicColumn,
    epicColumn: epicColumn
  });
  
  // Track results
  var requiredColumns = [];
  var recommendedColumns = [];
  var missingRequired = [];
  var missingRecommended = [];
  var foundColumns = [];
  
  // 1. Essential Jira columns - these are always needed
  requiredColumns.push({
    name: 'Ticket Key',
    purpose: 'Stores the created Jira ticket key (e.g., PROJ-123)',
    required: true
  });
  
  requiredColumns.push({
    name: 'Status',
    purpose: 'Stores the ticket status (Created, In Progress, etc.)',
    required: true
  });
  
  // 2. If using Epic, check for Epic column
  if (!tabSettings.jiraProject || tabSettings.jiraProject.trim() === '') {
    requiredColumns.push({
      name: 'EPIC',
      purpose: 'Required to determine project key when no project is set in tab settings',
      required: true
    });
  } else {
    recommendedColumns.push({
      name: 'EPIC',
      purpose: 'Used to associate tickets with specific Epics (optional)',
      required: false
    });
  }
  
  // 3. Check for '#' column for ticket link
  recommendedColumns.push({
    name: '#',
    purpose: 'Used to display hyperlinks to Jira tickets',
    required: false
  });
  
  // Check if there's any column mapped to Ticket ID
  var hasTicketIdMapping = false;
  for (var column in columnMapping) {
    if (columnMapping[column] && 
        ((typeof columnMapping[column] === 'object' && columnMapping[column].mapped === 'ticketId') ||
        columnMapping[column] === 'ticketId')) {
      hasTicketIdMapping = true;
      break;
    }
  }
  
  // If no ticketId mapping, recommend it
  if (!hasTicketIdMapping) {
    recommendedColumns.push({
      name: 'Any column mapped to "Jira Ticket ID"',
      purpose: 'Use a column to store ticket ID hyperlinks',
      required: false
    });
  }
  
  // 4. Check mapped columns that are marked as required
  for (var column in columnMapping) {
    if (columnMapping[column] && 
        typeof columnMapping[column] === 'object' && 
        columnMapping[column].mapped && 
        columnMapping[column].required) {
      
      requiredColumns.push({
        name: column,
        purpose: 'Mapped to ' + columnMapping[column].mapped + ' (required in mapping)',
        required: true
      });
    }
  }
  
  // 5. Check for Summary column if not mapped elsewhere
  var hasSummaryMapping = false;
  for (var column in columnMapping) {
    if (columnMapping[column] && 
        ((typeof columnMapping[column] === 'object' && columnMapping[column].mapped === 'summary') ||
         columnMapping[column] === 'summary')) {
      hasSummaryMapping = true;
      break;
    }
  }
  
  if (!hasSummaryMapping) {
    requiredColumns.push({
      name: 'Summary',
      purpose: 'Required for Jira ticket summary (no mapping found)',
      required: true
    });
  }
  
  // Check which columns are present/missing
  headers.forEach(function(header) {
    if (header) { // Skip empty headers
      foundColumns.push(header);
    }
  });
  
  // Find missing required columns
  requiredColumns.forEach(function(colInfo) {
    if (foundColumns.indexOf(colInfo.name) === -1 && 
        foundColumns.indexOf(colInfo.name.toLowerCase()) === -1 && 
        foundColumns.indexOf(colInfo.name.toUpperCase()) === -1) {
      missingRequired.push(colInfo);
    }
  });
  
  // Find missing recommended columns
  recommendedColumns.forEach(function(colInfo) {
    if (foundColumns.indexOf(colInfo.name) === -1 && 
        foundColumns.indexOf(colInfo.name.toLowerCase()) === -1 && 
        foundColumns.indexOf(colInfo.name.toUpperCase()) === -1) {
      missingRecommended.push(colInfo);
    }
  });
  
  // Log results
  logToSheet('analyzeSheetColumns', 'info', 'Column analysis results', {
    foundColumns: foundColumns,
    missingRequired: missingRequired,
    missingRecommended: missingRecommended
  });
  
  // Create HTML output with results
  var html = '<html><body style="font-family: Arial, sans-serif; padding: 20px;">';
  html += '<h2>Sheet Column Analysis: ' + sheet.getName() + '</h2>';
  
  // Show missing required columns with high importance
  if (missingRequired.length > 0) {
    html += '<h3 style="color: red;">Missing Required Columns</h3>';
    html += '<p>The following columns are required but missing from your sheet:</p>';
    html += '<table border="1" cellpadding="4" cellspacing="0" style="border-collapse: collapse;">';
    html += '<tr style="background-color: #f2f2f2;"><th>Column Name</th><th>Purpose</th></tr>';
    
    missingRequired.forEach(function(col) {
      html += '<tr style="background-color: #ffeeee;">';
      html += '<td><b>' + col.name + '</b></td>';
      html += '<td>' + col.purpose + '</td>';
      html += '</tr>';
    });
    
    html += '</table>';
  } else {
    html += '<p style="color: green;"><b>✓ All required columns are present in your sheet.</b></p>';
  }
  
  // Show missing recommended columns
  if (missingRecommended.length > 0) {
    html += '<h3 style="color: orange;">Missing Recommended Columns</h3>';
    html += '<p>The following columns are recommended but not required:</p>';
    html += '<table border="1" cellpadding="4" cellspacing="0" style="border-collapse: collapse;">';
    html += '<tr style="background-color: #f2f2f2;"><th>Column Name</th><th>Purpose</th></tr>';
    
    missingRecommended.forEach(function(col) {
      html += '<tr style="background-color: #fffbe5;">';
      html += '<td><b>' + col.name + '</b></td>';
      html += '<td>' + col.purpose + '</td>';
      html += '</tr>';
    });
    
    html += '</table>';
  }
  
  // Show column mapping summary
  html += '<h3>Column Mapping Summary</h3>';
  var mappedColumns = 0;
  
  for (var column in columnMapping) {
    if (columnMapping[column]) mappedColumns++;
  }
  
  if (mappedColumns === 0) {
    html += '<p style="color: red;"><b>Warning:</b> No column mappings are configured. ' + 
            'Use the "Configure Column Mapping" option to set up column mappings.</p>';
  } else {
    html += '<p>' + mappedColumns + ' column(s) mapped to Jira fields. ' +
            '<a href="#" onclick="google.script.run.viewColumnMapping(); return false;">View details</a></p>';
  }
  
  // Show tab settings summary
  html += '<h3>Tab Settings Summary</h3>';
  html += '<p>';
  
  if (tabSettings.jiraProject) {
    html += '<b>Project Key:</b> ' + tabSettings.jiraProject + '<br>';
  } else {
    html += '<span style="color: ' + (hasEpicColumn ? 'green' : 'red') + ';">' + 
            '<b>Project Key:</b> Not set in tab settings' + 
            (hasEpicColumn ? ' (will use EPIC column for project determination)' : ' (requires EPIC column for project determination)') + 
            '</span><br>';
  }
  
  html += '<b>Default Issue Type:</b> ' + (tabSettings.defaultIssueType || 'Task') + '<br>';
  
  if (tabSettings.epicFieldId) {
    html += '<b>Epic Field ID:</b> ' + tabSettings.epicFieldId;
  } else {
    html += '<span style="color: ' + (hasEpicColumn ? 'red' : 'gray') + ';">' + 
            '<b>Epic Field ID:</b> Not set ' + 
            (hasEpicColumn ? '(required for Epic linking)' : '(only needed if using EPIC column)') + 
            '</span>';
  }
  
  html += '</p>';
  
  // Add instructions
  html += '<h3>Next Steps</h3>';
  if (missingRequired.length > 0) {
    html += '<ol>';
    html += '<li>Add the missing required columns to row 2 of your sheet</li>';
    html += '<li>Configure column mappings using the "Configure Column Mapping" option</li>';
    if (!tabSettings.jiraProject && !hasEpicColumn) {
      html += '<li>Either configure a Project Key in tab settings or add an EPIC column</li>';
    }
    html += '</ol>';
  } else if (mappedColumns === 0) {
    html += '<p>Configure column mappings using the "Configure Column Mapping" option.</p>';
  } else {
    html += '<p>Your sheet appears to have all necessary columns. Use the "Test Row Validation" option to verify specific rows.</p>';
  }
  
  html += '</body></html>';
  
  // Show the dialog
  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(500)
    .setTitle('Sheet Column Analysis');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Sheet Column Analysis');
}

/**
 * Finds the column index (1-based) of a column that has a specific mapping or name
 * 
 * @since TBD
 * @param {Array}  headers          Array of header values from the sheet.
 * @param {string} mappingType      The mapping type to search for (e.g., 'ticketId', 'summary').
 * @param {string|Array} literalName Optional literal column name(s) to check for first.
 * @param {Object} columnMapping    The column mapping configuration (optional).
 * @return {number} The 1-based column index, or 0 if not found.
 */
function findColumnByMapping(headers, mappingType, literalName, columnMapping) {
  // Check for literal column name first if provided
  if (literalName) {
    // Handle both string and array of strings
    var literalNames = Array.isArray(literalName) ? literalName : [literalName];
    
    for (var i = 0; i < literalNames.length; i++) {
      var literalColumn = headers.indexOf(literalNames[i]) + 1;
      
      // If found, return it
      if (literalColumn > 0) {
        return literalColumn;
      }
    }
  }
  
  // If no mapping provided, get it
  if (!columnMapping) {
    columnMapping = getColumnMapping();
  }
  
  // Look for any column mapped to the specified type
  for (var column in columnMapping) {
    if (columnMapping[column] && headers.indexOf(column) >= 0) {
      var mapping = columnMapping[column];
      if ((typeof mapping === 'object' && mapping.mapped === mappingType) || 
          mapping === mappingType) {
        return headers.indexOf(column) + 1;
      }
    }
  }
  
  // No matching column found
  return 0;
}

/**
 * Finds the column index (1-based) of a column that stores ticket keys
 * Checks for both a literal "Ticket Key" column and columns mapped to ticketId
 *
 * @since TBD
 * @param {Array}  headers       Array of header values from the sheet.
 * @param {Object} columnMapping The column mapping configuration (optional).
 * @return {number} The 1-based column index, or 0 if not found.
 */
function findTicketKeyColumn(headers, columnMapping) {
  return findColumnByMapping(headers, 'ticketId', 'Ticket Key', columnMapping);
} 