# Google Sheets to Jira Integration

This Google Apps Script allows you to integrate your Google Sheets with Jira, making it easy to create and manage Jira tickets directly from your spreadsheet data.

## Features

- Configure Jira connection settings
- Map spreadsheet columns to Jira fields
- Create Jira tickets from selected data
- Bulk create tickets from entire spreadsheet
- Create tickets using predefined templates
- Automatically associate tickets with EPICs
- Sync ticket status from Jira back to the spreadsheet

## Setup Instructions

### 1. Copy the Script Files

1. Open your Google Sheets document
2. Go to **Extensions** > **Apps Script**
3. Create the following files in the Apps Script editor:
   - `Code.gs`
   - `Settings.html`
   - `ColumnMapping.html`
   - `TicketTemplates.gs`
   - `TicketTemplate.html`
4. Copy and paste the provided code into each file
5. Save the project with a name (e.g. "Sheets to Jira")

### 2. Configure Your Jira API

Before using the integration, you'll need to:

1. Create a Jira API token:
   - Log in to your Atlassian account
   - Go to **Account Settings** > **Security** > **Create and manage API tokens**
   - Create a new API token and copy it

2. In Google Sheets, refresh the page to see the new "Jira Integration" menu
3. Click **Jira Integration** > **Configure Settings**
4. Enter your Jira information:
   - **Jira URL**: Your Jira instance URL (e.g., `https://your-domain.atlassian.net`)
   - **Email**: Your Jira account email
   - **API Token**: The API token you created
   - **Project Key**: The key of your Jira project (e.g., `PROJ`)
   - **Default Issue Type**: The default issue type (e.g., `Task`, `Bug`, or `Story`)
   - **Epic Link Field ID**: The custom field ID for Epic Link (usually `customfield_10014`)
5. Click **Save Settings**
6. Test your connection using **Jira Integration** > **Test Connection**

## Usage Instructions

### Preparing Your Spreadsheet

Your spreadsheet should include columns that map to Jira fields. Recommended columns:

- **Summary**: The title of the Jira ticket
- **Description**: Detailed description of the ticket
- **Issue Type**: Type of issue (Task, Bug, Story, etc.)
- **Priority**: Priority level (Highest, High, Medium, Low, Lowest)
- **Labels**: Comma-separated list of labels
- **Epic Link**: Key of the parent Epic (e.g., `PROJ-123`)
- **Components**: Comma-separated list of components
- **Ticket Key**: Will be filled with the created ticket key
- **Status**: Will be filled with the ticket status

### Configuring Column Mapping

1. Click **Jira Integration** > **Configure Column Mapping**
2. For each column in your spreadsheet, select the corresponding Jira field
3. Click **Save Mapping**

### Creating Tickets

#### From Selected Rows:

1. Select the rows containing the data for new tickets
2. Click **Jira Integration** > **Create Jira Tickets**
3. The created ticket keys will appear in the "Ticket Key" column

#### Bulk Creation:

1. Ensure your spreadsheet has all required data
2. Click **Jira Integration** > **Bulk Create Tickets**
3. Confirm the bulk creation
4. The created ticket keys will appear in the "Ticket Key" column

#### Using Templates:

1. Click **Jira Integration** > **Create Ticket from Template**
2. Select a template (Bug, Task, or Story)
3. Fill in the required fields
4. Click **Create Ticket**

### Syncing Ticket Status

1. Ensure your spreadsheet has "Ticket Key" and "Status" columns
2. Click **Jira Integration** > **Sync Ticket Status**
3. The script will update the status of all tickets in the sheet

## Troubleshooting

- If you see "Please configure Jira settings first!" make sure you've completed the setup process
- If connection fails, check your Jira URL, email, and API token
- Ensure your spreadsheet has the necessary columns for creating tickets
- For Epic linking issues, verify your Epic Link Field ID is correct

## Customization

You can modify the script to:
- Add support for additional Jira fields
- Customize ticket templates
- Add additional validation rules
- Implement more advanced synchronization

## Limitations

- The script requires a Jira API token, which may need to be refreshed periodically
- Custom Jira fields require you to know the exact custom field ID
- Large bulk operations may time out due to Google Apps Script execution limits

## Security Notes

- Your Jira API token is stored in the Google Apps Script properties and is only accessible to your script
- Consider limiting sharing of the spreadsheet to control who can create tickets
- Review Google Apps Script permissions when first running the script 