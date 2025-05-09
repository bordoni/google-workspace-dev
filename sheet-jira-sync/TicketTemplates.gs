/**
 * Ticket Templates Utility
 * 
 * This file contains functions for template-based Jira ticket creation
 */

/**
 * Get predefined ticket templates from script properties
 */
function getTicketTemplates() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var templatesStr = scriptProperties.getProperty('ticketTemplates');
  
  if (templatesStr) {
    return JSON.parse(templatesStr);
  }
  
  // Return default templates if none are saved
  return {
    'Bug': {
      issuetype: { name: 'Bug' },
      summary: '[BUG] ${Summary}',
      description: "h2. Description\n${Description}\n\nh2. Steps to Reproduce\n${Steps to Reproduce}\n\nh2. Expected Result\n${Expected Result}\n\nh2. Actual Result\n${Actual Result}"
    },
    'Task': {
      issuetype: { name: 'Task' },
      summary: '${Summary}',
      description: "h2. Description\n${Description}\n\nh2. Acceptance Criteria\n${Acceptance Criteria}"
    },
    'Story': {
      issuetype: { name: 'Story' },
      summary: '${Summary}',
      description: "h2. Description\n${Description}\n\nh2. Acceptance Criteria\n${Acceptance Criteria}\n\nh2. Value\n${Value}"
    }
  };
}

/**
 * Save ticket templates to script properties
 */
function saveTicketTemplates(templates) {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('ticketTemplates', JSON.stringify(templates));
  return 'Ticket templates saved successfully!';
}

/**
 * Show a dialog to create a ticket from a template
 */
function showCreateFromTemplateDialog() {
  var html = HtmlService.createHtmlOutputFromFile('TicketTemplate')
    .setWidth(600)
    .setHeight(600)
    .setTitle('Create Jira Ticket from Template');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Create Jira Ticket from Template');
}

/**
 * Apply a template to the provided data
 * 
 * @param {string} templateName - Name of the template to apply
 * @param {object} fieldValues - Values for template variables
 * @return {object} Ticket data ready for creation
 */
function applyTemplate(templateName, fieldValues) {
  var templates = getTicketTemplates();
  
  if (!templates.hasOwnProperty(templateName)) {
    throw new Error('Template not found: ' + templateName);
  }
  
  var template = templates[templateName];
  var settings = getSettings();
  
  // Create a copy of the template to modify
  var ticketData = {
    fields: {
      project: {
        key: settings.jiraProject
      }
    }
  };
  
  // Copy template fields
  Object.keys(template).forEach(function(field) {
    if (typeof template[field] === 'string') {
      // Apply variable substitution for string fields
      ticketData.fields[field] = substituteVariables(template[field], fieldValues);
    } else {
      // Copy non-string fields directly
      ticketData.fields[field] = template[field];
    }
  });
  
  // Add Epic link if configured and present in the field values
  if (settings.epicFieldId && fieldValues.hasOwnProperty('EpicLink')) {
    ticketData.fields[settings.epicFieldId] = fieldValues['EpicLink'];
  }
  
  return ticketData;
}

/**
 * Substitute variables in a template string
 * 
 * @param {string} templateString - The template string with ${variable} placeholders
 * @param {object} values - Object containing the variable values
 * @return {string} The string with variables replaced
 */
function substituteVariables(templateString, values) {
  // Replace ${variable} with its value from the values object
  return templateString.replace(/\${([^}]+)}/g, function(match, variable) {
    return values.hasOwnProperty(variable) ? values[variable] : match;
  });
}

/**
 * Create a ticket from a template
 * 
 * @param {string} templateName - Name of the template to use
 * @param {object} fieldValues - Values for template variables
 * @return {object} Created ticket details
 */
function createTicketFromTemplate(templateName, fieldValues) {
  var ticketData = applyTemplate(templateName, fieldValues);
  var response = createJiraTicket(ticketData);
  
  if (response.getResponseCode() === 201) {
    return JSON.parse(response.getContentText());
  } else {
    throw new Error('Failed to create ticket: ' + response.getContentText());
  }
}

/**
 * Create tickets for selected rows using a template
 * 
 * @param {string} templateName - Name of the template to use
 * @return {object} Result of the operation
 */
function createTicketsFromTemplateForSelection(templateName) {
  var settings = getSettings();
  
  if (!settings.jiraUrl || !settings.jiraEmail || !settings.jiraApiToken || !settings.jiraProject) {
    throw new Error('Jira settings not configured');
  }
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var selectedRange = SpreadsheetApp.getActiveRange();
  
  if (!selectedRange) {
    throw new Error('No data selected');
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
    
    // Skip rows that already have a ticket key
    if (ticketKeyColumn > 0 && sheet.getRange(rowIndex, ticketKeyColumn).getValue()) {
      continue;
    }
    
    try {
      // Convert row data to field values for the template
      var fieldValues = {};
      for (var j = 0; j < headers.length; j++) {
        if (rowData[j]) {
          fieldValues[headers[j]] = rowData[j];
        }
      }
      
      // Apply template and create ticket
      var ticketData = applyTemplate(templateName, fieldValues);
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
  
  return {
    created: createdCount,
    errors: errorCount,
    errorMessages: errorMessages
  };
} 