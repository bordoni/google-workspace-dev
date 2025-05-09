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
    .addItem('Create Jira Tickets', 'createJiraTicketsFromSelection')
    .addItem('Bulk Create Tickets', 'createJiraTicketsFromSheet')
    .addItem('Create Ticket from Template', 'showCreateFromTemplateDialog')
    .addSeparator()
    .addItem('Sync Ticket Status', 'syncJiraStatus')
    .addSeparator()
    .addItem('Test Connection', 'testJiraConnection')
    .addToUi();
}

/**
 * Display a dialog for configuring Jira API settings
 */
function showSettingsDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Settings')
    .setWidth(600)
    .setHeight(500)
    .setTitle('Jira Integration Settings');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Jira Integration Settings');
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
  
  scriptProperties.setProperties({
    'jiraUrl': settings.jiraUrl,
    'jiraEmail': settings.jiraEmail,
    'jiraApiToken': settings.jiraApiToken,
    'jiraProject': settings.jiraProject,
    'defaultIssueType': settings.defaultIssueType,
    'epicFieldId': settings.epicFieldId
  });
  
  return 'Settings saved successfully!';
}

/**
 * Get the current Jira settings
 */
function getSettings() {
  var scriptProperties = PropertiesService.getScriptProperties();
  return {
    jiraUrl: scriptProperties.getProperty('jiraUrl'),
    jiraEmail: scriptProperties.getProperty('jiraEmail'),
    jiraApiToken: scriptProperties.getProperty('jiraApiToken'),
    jiraProject: scriptProperties.getProperty('jiraProject'),
    defaultIssueType: scriptProperties.getProperty('defaultIssueType'),
    epicFieldId: scriptProperties.getProperty('epicFieldId')
  };
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
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
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
 * Create Jira tickets from selected sheet data
 */
function createJiraTicketsFromSelection() {
  var settings = getSettings();
  
  if (!settings.jiraUrl || !settings.jiraEmail || !settings.jiraApiToken || !settings.jiraProject) {
    SpreadsheetApp.getUi().alert('Please configure Jira settings first!');
    return;
  }
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var selectedRange = SpreadsheetApp.getActiveRange();
  
  if (!selectedRange) {
    SpreadsheetApp.getUi().alert('Please select data rows to create tickets');
    return;
  }
  
  // Get column headers
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Process each selected row
  var dataRows = selectedRange.getValues();
  var statusColumn = headers.indexOf('Status') + 1;
  var ticketKeyColumn = headers.indexOf('Ticket Key') + 1;
  
  var createdCount = 0;
  var errorCount = 0;
  var errorMessages = [];
  
  for (var i = 0; i < dataRows.length; i++) {
    var rowData = dataRows[i];
    var rowIndex = selectedRange.getRowIndex() + i;
    
    // Skip rows that already have a ticket key or are marked as processed
    if (ticketKeyColumn > 0 && sheet.getRange(rowIndex, ticketKeyColumn).getValue()) {
      continue;
    }
    
    try {
      var ticketData = prepareTicketData(headers, rowData);
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
  }
  
  var message = createdCount + ' ticket(s) created successfully.\n';
  if (errorCount > 0) {
    message += errorCount + ' error(s) occurred:\n' + errorMessages.join('\n');
  }
  
  SpreadsheetApp.getUi().alert(message);
}

/**
 * Create Jira tickets from the entire sheet (bulk creation)
 */
function createJiraTicketsFromSheet() {
  var settings = getSettings();
  
  if (!settings.jiraUrl || !settings.jiraEmail || !settings.jiraApiToken || !settings.jiraProject) {
    SpreadsheetApp.getUi().alert('Please configure Jira settings first!');
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
  var statusColumn = headers.indexOf('Status') + 1;
  var ticketKeyColumn = headers.indexOf('Ticket Key') + 1;
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
 * Prepare ticket data from row data using the fixed mapping
 */
function prepareTicketData(headers, rowData) {
  var settings = getSettings();
  var ticketData = {
    fields: {
      project: {
        key: settings.jiraProject
      },
      issuetype: {
        name: settings.defaultIssueType
      }
    }
  };
  
  // Map sheet columns to Jira fields
  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    var value = rowData[i];
    
    if (!value) continue;
    
    // Handle special fields
    switch (header) {
      case 'Summary':
        ticketData.fields.summary = value;
        break;
        
      case 'Description':
        ticketData.fields.description = value;
        break;
        
      case 'Issue Type':
        ticketData.fields.issuetype.name = value;
        break;
        
      case 'Priority':
        ticketData.fields.priority = { name: value };
        break;
        
      case 'Labels':
        // Split comma-separated labels
        if (typeof value === 'string') {
          ticketData.fields.labels = value.split(',').map(function(label) {
            return label.trim();
          });
        }
        break;
        
      case 'Epic Link':
        // If EPIC field ID is configured and value exists, link to EPIC
        if (settings.epicFieldId && value) {
          ticketData.fields[settings.epicFieldId] = value;
        }
        break;
        
      case 'Components':
        // Handle multiple components as comma-separated values
        if (typeof value === 'string') {
          var components = value.split(',').map(function(component) {
            return { name: component.trim() };
          });
          ticketData.fields.components = components;
        }
        break;
        
      // Add more field mappings as needed
    }
  }
  
  return ticketData;
}

/**
 * Prepare ticket data from row data using the configured column mapping
 */
function prepareTicketDataWithMapping(headers, rowData) {
  var settings = getSettings();
  var columnMapping = getColumnMapping();
  
  var ticketData = {
    fields: {
      project: {
        key: settings.jiraProject
      },
      issuetype: {
        name: settings.defaultIssueType
      }
    }
  };
  
  // Map sheet columns to Jira fields using the configured mapping
  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    var value = rowData[i];
    
    if (!value || !columnMapping[header]) continue;
    
    var jiraField = columnMapping[header];
    
    // Handle different types of fields
    switch (jiraField) {
      case 'summary':
        ticketData.fields.summary = value;
        break;
        
      case 'description':
        ticketData.fields.description = value;
        break;
        
      case 'issuetype':
        ticketData.fields.issuetype.name = value;
        break;
        
      case 'priority':
        ticketData.fields.priority = { name: value };
        break;
        
      case 'labels':
        // Split comma-separated labels
        if (typeof value === 'string') {
          ticketData.fields.labels = value.split(',').map(function(label) {
            return label.trim();
          });
        } else if (Array.isArray(value)) {
          ticketData.fields.labels = value;
        }
        break;
        
      case 'epicLink':
        // If EPIC field ID is configured and value exists, link to EPIC
        if (settings.epicFieldId && value) {
          ticketData.fields[settings.epicFieldId] = value;
        }
        break;
        
      case 'components':
        // Handle multiple components as comma-separated values
        if (typeof value === 'string') {
          var components = value.split(',').map(function(component) {
            return { name: component.trim() };
          });
          ticketData.fields.components = components;
        }
        break;
        
      case 'assignee':
        ticketData.fields.assignee = { name: value };
        break;
        
      case 'reporter':
        ticketData.fields.reporter = { name: value };
        break;
        
      case 'duedate':
        // Format date as YYYY-MM-DD if it's a Date object
        if (value instanceof Date) {
          ticketData.fields.duedate = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else {
          ticketData.fields.duedate = value;
        }
        break;
    }
  }
  
  // Check if required fields are present
  if (!ticketData.fields.summary) {
    throw new Error('Summary field is required');
  }
  
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
  var settings = getSettings();
  
  if (!settings.jiraUrl || !settings.jiraEmail || !settings.jiraApiToken) {
    SpreadsheetApp.getUi().alert('Please configure Jira settings first!');
    return;
  }
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var ticketKeyColumn = headers.indexOf('Ticket Key') + 1;
  var statusColumn = headers.indexOf('Status') + 1;
  
  if (ticketKeyColumn < 1 || statusColumn < 1) {
    SpreadsheetApp.getUi().alert('Sheet must contain "Ticket Key" and "Status" columns!');
    return;
  }
  
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; // Only header row exists
  
  var ticketKeys = sheet.getRange(2, ticketKeyColumn, lastRow - 1, 1).getValues();
  var updatedCount = 0;
  var errorCount = 0;
  
  for (var i = 0; i < ticketKeys.length; i++) {
    var ticketKey = ticketKeys[i][0];
    if (!ticketKey) continue;
    
    try {
      var response = makeJiraRequest('issue/' + ticketKey, 'GET');
      if (response.getResponseCode() === 200) {
        var issueData = JSON.parse(response.getContentText());
        var status = issueData.fields.status.name;
        
        sheet.getRange(i + 2, statusColumn).setValue(status);
        updatedCount++;
      } else {
        errorCount++;
      }
    } catch (error) {
      errorCount++;
    }
  }
  
  SpreadsheetApp.getUi().alert(updatedCount + ' ticket status(es) updated.\n' + 
                              errorCount + ' error(s) occurred.');
} 