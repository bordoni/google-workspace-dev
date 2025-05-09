/**
 * Jira API Service
 * 
 * This file contains functions for communicating with the Jira API
 */

/**
 * Creates a Jira issue based on provided data
 * 
 * @param {Object} issueData - The data to create the issue with
 * @return {Object} Response from Jira API containing the created issue
 */
function createJiraIssue(issueData) {
  const settings = getJiraSettings();
  if (!settings) {
    throw new Error('Jira settings not configured');
  }
  
  const url = `${settings.jiraUrl}/rest/api/2/issue`;
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(`${settings.email}:${settings.apiToken}`)
    },
    payload: JSON.stringify(issueData),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  
  if (responseCode < 200 || responseCode >= 300) {
    Logger.log(`Error creating issue: ${responseText}`);
    throw new Error(`Failed to create Jira issue: ${responseText}`);
  }
  
  return JSON.parse(responseText);
}

/**
 * Tests the connection to Jira API
 * 
 * @return {Boolean} Whether the connection was successful
 */
function testJiraApiConnection() {
  const settings = getJiraSettings();
  if (!settings) {
    return false;
  }
  
  try {
    const url = `${settings.jiraUrl}/rest/api/2/myself`;
    const options = {
      method: 'get',
      headers: {
        'Authorization': 'Basic ' + Utilities.base64Encode(`${settings.email}:${settings.apiToken}`)
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    return responseCode >= 200 && responseCode < 300;
  } catch (error) {
    Logger.log(`Error testing connection: ${error.message}`);
    return false;
  }
}

/**
 * Gets an issue from Jira by its key
 * 
 * @param {String} issueKey - The key of the issue to get
 * @return {Object} The issue data
 */
function getJiraIssue(issueKey) {
  const settings = getJiraSettings();
  if (!settings) {
    throw new Error('Jira settings not configured');
  }
  
  const url = `${settings.jiraUrl}/rest/api/2/issue/${issueKey}`;
  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(`${settings.email}:${settings.apiToken}`)
    },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  
  if (responseCode < 200 || responseCode >= 300) {
    Logger.log(`Error getting issue: ${response.getContentText()}`);
    throw new Error(`Failed to get Jira issue ${issueKey}`);
  }
  
  return JSON.parse(response.getContentText());
}

/**
 * Gets all issue types for the configured Jira project
 * 
 * @return {Array} Array of issue types
 */
function getJiraIssueTypes() {
  const settings = getJiraSettings();
  if (!settings) {
    throw new Error('Jira settings not configured');
  }
  
  const url = `${settings.jiraUrl}/rest/api/2/issue/createmeta?projectKeys=${settings.projectKey}&expand=projects.issuetypes`;
  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(`${settings.email}:${settings.apiToken}`)
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode < 200 || responseCode >= 300) {
      Logger.log(`Error getting issue types: ${response.getContentText()}`);
      throw new Error('Failed to get Jira issue types');
    }
    
    const responseData = JSON.parse(response.getContentText());
    if (!responseData.projects || responseData.projects.length === 0) {
      return [];
    }
    
    return responseData.projects[0].issuetypes || [];
  } catch (error) {
    Logger.log(`Error getting issue types: ${error.message}`);
    return [];
  }
}

/**
 * Gets all field metadata for a specific issue type
 * 
 * @param {String} issueTypeId - The ID of the issue type
 * @return {Array} Array of field metadata
 */
function getFieldsForIssueType(issueTypeId) {
  const settings = getJiraSettings();
  if (!settings) {
    throw new Error('Jira settings not configured');
  }
  
  const url = `${settings.jiraUrl}/rest/api/2/issue/createmeta?projectKeys=${settings.projectKey}&issuetypeIds=${issueTypeId}&expand=projects.issuetypes.fields`;
  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(`${settings.email}:${settings.apiToken}`)
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode < 200 || responseCode >= 300) {
      Logger.log(`Error getting fields: ${response.getContentText()}`);
      throw new Error('Failed to get Jira fields');
    }
    
    const responseData = JSON.parse(response.getContentText());
    if (!responseData.projects || responseData.projects.length === 0 || 
        !responseData.projects[0].issuetypes || responseData.projects[0].issuetypes.length === 0) {
      return {};
    }
    
    return responseData.projects[0].issuetypes[0].fields || {};
  } catch (error) {
    Logger.log(`Error getting fields: ${error.message}`);
    return {};
  }
}

/**
 * Gets all available fields from Jira
 * 
 * @return {Array} Array of available fields
 */
function getAllJiraFields() {
  const settings = getJiraSettings();
  if (!settings) {
    throw new Error('Jira settings not configured');
  }
  
  const url = `${settings.jiraUrl}/rest/api/2/field`;
  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(`${settings.email}:${settings.apiToken}`)
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode < 200 || responseCode >= 300) {
      Logger.log(`Error getting all fields: ${response.getContentText()}`);
      throw new Error('Failed to get Jira fields');
    }
    
    return JSON.parse(response.getContentText()) || [];
  } catch (error) {
    Logger.log(`Error getting all fields: ${error.message}`);
    return [];
  }
}

/**
 * Gets all priorities from Jira
 * 
 * @return {Array} Array of priorities
 */
function getJiraPriorities() {
  const settings = getJiraSettings();
  if (!settings) {
    throw new Error('Jira settings not configured');
  }
  
  const url = `${settings.jiraUrl}/rest/api/2/priority`;
  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(`${settings.email}:${settings.apiToken}`)
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode < 200 || responseCode >= 300) {
      Logger.log(`Error getting priorities: ${response.getContentText()}`);
      throw new Error('Failed to get Jira priorities');
    }
    
    return JSON.parse(response.getContentText()) || [];
  } catch (error) {
    Logger.log(`Error getting priorities: ${error.message}`);
    return [];
  }
}

/**
 * Gets all components for the configured project
 * 
 * @return {Array} Array of components
 */
function getProjectComponents() {
  const settings = getJiraSettings();
  if (!settings) {
    throw new Error('Jira settings not configured');
  }
  
  const url = `${settings.jiraUrl}/rest/api/2/project/${settings.projectKey}/components`;
  const options = {
    method: 'get',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(`${settings.email}:${settings.apiToken}`)
    },
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode < 200 || responseCode >= 300) {
      Logger.log(`Error getting components: ${response.getContentText()}`);
      throw new Error('Failed to get Jira components');
    }
    
    return JSON.parse(response.getContentText()) || [];
  } catch (error) {
    Logger.log(`Error getting components: ${error.message}`);
    return [];
  }
}

/**
 * Builds a properly formatted Jira issue create request
 * 
 * @param {Object} data - The issue data
 * @param {String} epicKey - Optional epic key to link to
 * @return {Object} Properly formatted issue create request
 */
function buildJiraIssueRequest(data, epicKey) {
  const settings = getJiraSettings();
  if (!settings) {
    throw new Error('Jira settings not configured');
  }
  
  // Start with the basic issue structure
  const issueRequest = {
    fields: {
      project: {
        key: settings.projectKey
      },
      summary: data.summary || 'No summary provided',
      issuetype: {
        name: data.issuetype || settings.defaultIssueType || 'Task'
      }
    }
  };
  
  // Add description if provided
  if (data.description) {
    issueRequest.fields.description = data.description;
  }
  
  // Add priority if provided
  if (data.priority) {
    issueRequest.fields.priority = {
      name: data.priority
    };
  }
  
  // Add components if provided
  if (data.components && data.components.length > 0) {
    issueRequest.fields.components = data.components.map(component => ({
      name: component.trim()
    }));
  }
  
  // Add labels if provided
  if (data.labels && data.labels.length > 0) {
    issueRequest.fields.labels = data.labels;
  }
  
  // Add epic link if provided and epic link field ID is configured
  if (epicKey && settings.epicLinkFieldId) {
    issueRequest.fields[settings.epicLinkFieldId] = epicKey;
  }
  
  // Add any custom fields that are provided
  if (data.customFields && Object.keys(data.customFields).length > 0) {
    for (const [fieldId, value] of Object.entries(data.customFields)) {
      issueRequest.fields[fieldId] = value;
    }
  }
  
  return issueRequest;
} 