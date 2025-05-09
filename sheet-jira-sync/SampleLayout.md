# Sample Spreadsheet Layout

Below is a recommended layout for your Google Sheet to work effectively with the Jira integration.

## Column Structure

| A | B | C | D | E | F | G | H | I |
|---|---|---|---|---|---|---|---|---|
| **Summary** | **Description** | **Issue Type** | **Priority** | **Labels** | **Epic Link** | **Components** | **Ticket Key** | **Status** |
| Fix login bug | Users unable to login on mobile app | Bug | High | mobile,login | PROJ-123 | Authentication | *auto-filled* | *auto-filled* |
| Update user documentation | Add new feature documentation | Task | Medium | documentation | PROJ-124 | Documentation |  |  |
| Implement search feature | Add search functionality to dashboard | Story | High | feature,search |  | Frontend,API |  |  |

## Column Descriptions

### Required Columns

1. **Summary (A)**: The title of the Jira ticket
   - Keep it concise but descriptive
   - This will become the Jira issue summary

2. **Description (B)**: Detailed description of the ticket
   - Can include formatting using Jira markup
   - Supports multi-line text

### Optional but Recommended Columns

3. **Issue Type (C)**: The type of issue
   - Common values: Bug, Task, Story, Epic
   - If left blank, the default issue type from settings will be used

4. **Priority (D)**: Priority level
   - Standard values: Highest, High, Medium, Low, Lowest
   - If left blank, the default priority (usually Medium) will be used

5. **Labels (E)**: Comma-separated list of labels
   - Format: label1,label2,label3
   - No spaces between commas

6. **Epic Link (F)**: Key of the parent Epic
   - Format: PROJ-123
   - Leave blank if the issue doesn't belong to an Epic

7. **Components (G)**: Comma-separated list of components
   - Format: component1,component2
   - Components must already exist in your Jira project

### Auto-filled Columns

8. **Ticket Key (H)**: Will be filled with the created ticket key
   - Format: PROJ-123
   - This column is used to track which tickets have been created
   - Used for status syncing

9. **Status (I)**: Will be filled with the ticket status
   - Updated when using the "Sync Ticket Status" feature
   - Examples: To Do, In Progress, Done

## Additional Tips

- Add headers to the first row of your spreadsheet
- Format the spreadsheet using Google Sheets' formatting tools for better readability
- You can freeze the header row (View > Freeze > 1 row)
- Use data validation for Issue Type and Priority columns to limit inputs to valid values
- If you need additional fields, add them as new columns and configure them in the Column Mapping dialog

## Example Use Cases

### Bug Tracking

| Summary | Description | Issue Type | Priority | Labels | Epic Link | Components | Ticket Key | Status |
|---|---|---|---|---|---|---|---|---|
| Login failure on Safari | Users report unable to login using Safari browser | Bug | High | browser,safari | PROJ-123 | Authentication |  |  |

### Feature Development

| Summary | Description | Issue Type | Priority | Labels | Epic Link | Components | Ticket Key | Status |
|---|---|---|---|---|---|---|---|---|
| Add dark mode support | Implement dark mode theme for mobile app | Story | Medium | ui,theme | PROJ-456 | UI |  |  |
| Create color palette | Design color system for dark mode | Task | Medium | design | PROJ-456 | Design |  |  |

### Support Tickets

| Summary | Description | Issue Type | Priority | Labels | Epic Link | Components | Ticket Key | Status |
|---|---|---|---|---|---|---|---|---|
| User unable to reset password | Customer reported issue via support chat | Support | Medium | customer,password |  | Authentication |  |  | 