# Microsoft MCP Server - Complete API Reference

## Overview

This document provides detailed information about all 10 tools available in the Microsoft MCP server, including parameters, responses, and usage examples.

---

## Email Tools

### 1. microsoft_search_emails

**Purpose**: Search emails with full-text search, folder filtering, and unread status.

**Parameters**:
```json
{
  "query": "string (required, 1-500 chars)",
  "folder": "string (default: inbox)",
  "limit": "integer (default: 20, 1-100)",
  "unread_only": "boolean (default: false)"
}
```

**Folder Options**:
- `inbox` - Inbox folder
- `sent` - Sent items
- `drafts` - Draft emails
- `archive` - Archive folder

**Response**:
```json
{
  "search_query": "budget report",
  "folder": "inbox",
  "results_count": 3,
  "emails": [
    {
      "id": "AAMkADE4...",
      "subject": "Q1 Budget Report",
      "from": "alice@example.com",
      "received_date": "2024-01-15T10:30:00Z",
      "unread": false,
      "importance": "normal",
      "preview": "Here's the Q1 budget breakdown..."
    }
  ]
}
```

**Examples**:
```
Claude: "Find emails about the project budget"
Params: {"query": "project budget", "folder": "inbox", "limit": 10}

Claude: "Show me unread emails from this week"
Params: {"unread_only": true, "folder": "inbox", "limit": 20}

Claude: "Search my sent items for emails to the CEO"
Params: {"query": "CEO", "folder": "sent", "limit": 5}
```

---

### 2. microsoft_send_email

**Purpose**: Send emails with HTML formatting, CC/BCC, and importance levels.

**Parameters**:
```json
{
  "to": ["email1@example.com"],  // required
  "cc": ["cc@example.com"],       // optional
  "bcc": ["bcc@example.com"],     // optional
  "subject": "string (required, 1-500 chars)",
  "body": "string (required, 1+ chars)",
  "importance": "low|normal|high (default: normal)",
  "is_html": "boolean (default: true)"
}
```

**Response**:
```json
{
  "status": "success",
  "message": "Email sent successfully",
  "recipients": ["alice@example.com"],
  "subject": "Project Update"
}
```

**Examples**:
```
Claude: "Send an email to john@example.com about the meeting"
Params: {
  "to": ["john@example.com"],
  "subject": "Meeting Scheduled",
  "body": "<p>The meeting is scheduled for 2 PM tomorrow.</p>",
  "importance": "normal"
}

Claude: "Send high priority email to multiple recipients"
Params: {
  "to": ["team@example.com"],
  "cc": ["manager@example.com"],
  "subject": "Urgent Action Required",
  "body": "<p><strong>Please respond by EOD.</strong></p>",
  "importance": "high",
  "is_html": true
}
```

---

### 3. microsoft_create_email_draft

**Purpose**: Create email drafts for later editing and sending.

**Parameters**:
```json
{
  "to": ["email@example.com"],      // optional
  "cc": ["cc@example.com"],         // optional
  "subject": "string (default: '')",
  "body": "string (required, 1+ chars)",
  "is_html": "boolean (default: true)"
}
```

**Response**:
```json
{
  "status": "success",
  "message": "Draft created successfully",
  "draft_id": "AAMkADE4...",
  "subject": "Project Proposal",
  "edit_link": "Draft ID: AAMkADE4..."
}
```

**Examples**:
```
Claude: "Create a draft email to the team"
Params: {
  "subject": "Weekly Team Update",
  "body": "<p>This week's highlights:</p><ul><li>Item 1</li></ul>",
  "to": ["team@example.com"]
}

Claude: "Draft a follow-up email"
Params: {
  "subject": "Re: Your inquiry",
  "body": "Thank you for reaching out. Here's the information you requested...",
  "is_html": true
}
```

---

### 4. microsoft_get_email_details

**Purpose**: Retrieve full email content, attachments, and metadata.

**Parameters**:
```json
{
  "message_id": "string (required, 1+ chars)",
  "include_attachments": "boolean (default: true)"
}
```

**Response**:
```json
{
  "id": "AAMkADE4...",
  "subject": "Project Update",
  "from": "alice@example.com",
  "to": ["bob@example.com"],
  "cc": ["manager@example.com"],
  "received_date": "2024-01-15T10:30:00Z",
  "sent_date": "2024-01-15T10:25:00Z",
  "unread": false,
  "importance": "normal",
  "body": "<p>Here's the project status...</p>",
  "categories": ["Work", "Important"],
  "attachments": [
    {
      "id": "AAAttachmentId",
      "name": "report.pdf",
      "size": 1024000
    }
  ]
}
```

**Examples**:
```
Claude: "Show me the full content of email ID AAMkADE4..."
Params: {
  "message_id": "AAMkADE4...",
  "include_attachments": true
}

Claude: "Get email details but don't show attachments"
Params: {
  "message_id": "AAMkADE4...",
  "include_attachments": false
}
```

---

### 5. microsoft_manage_email_folders

**Purpose**: List, create, delete, or rename email folders.

**Parameters**:
```json
{
  "action": "list|create|delete|rename (required)",
  "folder_name": "string (required for create/delete/rename)",
  "new_name": "string (required for rename)"
}
```

**Response for 'list' action**:
```json
{
  "action": "list",
  "folder_count": 5,
  "folders": [
    {
      "id": "AAMkFolder1",
      "name": "Important",
      "unread_count": 3,
      "subfolder_count": 2
    }
  ]
}
```

**Response for 'create' action**:
```json
{
  "action": "create",
  "status": "success",
  "folder_id": "AAMkNewFolder",
  "folder_name": "Projects"
}
```

**Examples**:
```
Claude: "List all my email folders"
Params: {"action": "list"}

Claude: "Create a folder called 'Client Work'"
Params: {"action": "create", "folder_name": "Client Work"}

Claude: "Rename folder 'Old Name' to 'New Name'"
Params: {"action": "rename", "folder_name": "Old Name", "new_name": "New Name"}

Claude: "Delete the 'Archive 2023' folder"
Params: {"action": "delete", "folder_name": "Archive 2023"}
```

---

## Calendar Tools

### 6. microsoft_create_calendar_event

**Purpose**: Create calendar events with attendees, location, and reminders.

**Parameters**:
```json
{
  "subject": "string (required, 1-500 chars)",
  "start_time": "string ISO 8601 (required, e.g., '2024-01-15T10:00:00Z')",
  "end_time": "string ISO 8601 (required, e.g., '2024-01-15T11:00:00Z')",
  "attendees": ["email@example.com"],  // optional
  "description": "string",             // optional
  "location": "string",                // optional
  "is_reminder": "boolean (default: true)"
}
```

**Response**:
```json
{
  "status": "success",
  "message": "Event created successfully",
  "event_id": "AAMkEvent123",
  "subject": "Team Meeting",
  "start_time": "2024-01-15T10:00:00Z",
  "end_time": "2024-01-15T11:00:00Z",
  "attendees_count": 3
}
```

**Examples**:
```
Claude: "Create a meeting with Alice and Bob next Tuesday at 2 PM"
Params: {
  "subject": "Project Kickoff",
  "start_time": "2024-01-16T14:00:00Z",
  "end_time": "2024-01-16T15:00:00Z",
  "attendees": ["alice@example.com", "bob@example.com"],
  "location": "Conference Room A",
  "description": "Initial project planning session"
}

Claude: "Block off time for lunch tomorrow"
Params: {
  "subject": "Lunch Break",
  "start_time": "2024-01-15T12:00:00Z",
  "end_time": "2024-01-15T13:00:00Z",
  "is_reminder": true
}
```

---

### 7. microsoft_check_availability

**Purpose**: Check calendar availability for a time slot.

**Parameters**:
```json
{
  "start_time": "string ISO 8601 (required, e.g., '2024-01-15T10:00:00Z')",
  "end_time": "string ISO 8601 (required, e.g., '2024-01-15T11:00:00Z')",
  "attendees": ["email@example.com"]  // optional
}
```

**Response**:
```json
{
  "requested_start": "2024-01-15T14:00:00Z",
  "requested_end": "2024-01-15T15:00:00Z",
  "is_available": true,
  "conflicts_count": 0,
  "conflicts": [],
  "attendees_checked": ["alice@example.com"],
  "note": "Individual attendee availability requires separate checks..."
}
```

**Examples**:
```
Claude: "Is 2 PM available tomorrow?"
Params: {
  "start_time": "2024-01-16T14:00:00Z",
  "end_time": "2024-01-16T15:00:00Z"
}

Claude: "Find slots that work for Alice and Bob next week"
Params: {
  "start_time": "2024-01-22T09:00:00Z",
  "end_time": "2024-01-22T17:00:00Z",
  "attendees": ["alice@example.com", "bob@example.com"]
}
```

---

### 8. microsoft_get_calendar_events

**Purpose**: Retrieve calendar events for a date range.

**Parameters**:
```json
{
  "start_date": "string YYYY-MM-DD (optional, default: today)",
  "end_date": "string YYYY-MM-DD (optional, default: today + 7 days)",
  "limit": "integer (default: 20, 1-100)",
  "include_details": "boolean (default: true)"
}
```

**Response**:
```json
{
  "date_range": "2024-01-15 to 2024-01-22",
  "event_count": 5,
  "events": [
    {
      "id": "AAMkEvent1",
      "subject": "Team Meeting",
      "start": "2024-01-15T10:00:00Z",
      "end": "2024-01-15T11:00:00Z",
      "is_organizer": true,
      "organizer": "you@example.com"
    }
  ]
}
```

**Examples**:
```
Claude: "What's on my calendar this week?"
Params: {
  "start_date": "2024-01-15",
  "end_date": "2024-01-21",
  "limit": 50
}

Claude: "Show my schedule for January"
Params: {
  "start_date": "2024-01-01",
  "end_date": "2024-01-31",
  "include_details": true
}
```

---

## Cloud Storage Tools

### 9. microsoft_list_cloud_files

**Purpose**: List files and folders on OneDrive, SharePoint, or Teams.

**Parameters**:
```json
{
  "location": "onedrive|sharepoint|teams (default: onedrive)",
  "folder_path": "string (default: '/')",
  "limit": "integer (default: 20, 1-100)",
  "recursive": "boolean (default: false)"
}
```

**Response**:
```json
{
  "location": "onedrive",
  "path": "/Projects",
  "item_count": 3,
  "items": [
    {
      "id": "01ABCDEF123456",
      "name": "Q1 Planning.xlsx",
      "type": "file",
      "size": 1048576,
      "modified": "2024-01-10T15:30:00Z",
      "web_url": "https://onedrive.live.com/..."
    }
  ]
}
```

**Examples**:
```
Claude: "Show me files in OneDrive /Projects folder"
Params: {
  "location": "onedrive",
  "folder_path": "/Projects",
  "limit": 20
}

Claude: "List SharePoint documents"
Params: {
  "location": "sharepoint",
  "folder_path": "/Shared Documents",
  "limit": 50
}

Claude: "What's in my Teams file storage?"
Params: {
  "location": "teams",
  "limit": 20
}
```

---

### 10. microsoft_create_folder

**Purpose**: Create new folders in cloud storage.

**Parameters**:
```json
{
  "folder_name": "string (required, 1+ chars)",
  "parent_path": "string (default: '/')",
  "location": "onedrive|sharepoint|teams (default: onedrive)"
}
```

**Response**:
```json
{
  "status": "success",
  "message": "Folder created successfully",
  "folder_id": "01NEWFOLDERID",
  "folder_name": "Q1 Planning",
  "web_url": "https://onedrive.live.com/..."
}
```

**Examples**:
```
Claude: "Create a folder called 'Q1 Projects' in OneDrive"
Params: {
  "folder_name": "Q1 Projects",
  "parent_path": "/",
  "location": "onedrive"
}

Claude: "Create a subfolder under /Projects/Active"
Params: {
  "folder_name": "Website Redesign",
  "parent_path": "/Projects/Active",
  "location": "onedrive"
}
```

---

## Response Codes & Error Handling

### Success Responses
- **200 OK**: Operation completed successfully
- Returns JSON with operation results

### Common Errors

#### Authentication Errors
```json
{
  "error": "Invalid client credentials",
  "status": "error"
}
```
→ Solution: Verify MICROSOFT_CLIENT_SECRET hasn't expired

#### Permission Errors
```json
{
  "error": "Permission denied to this resource",
  "status": "error"
}
```
→ Solution: Grant admin consent in Azure Portal

#### Rate Limit Errors
```json
{
  "error": "Rate limit exceeded, please try again later",
  "status": "error"
}
```
→ Solution: Wait a few seconds before retrying

#### Not Found Errors
```json
{
  "error": "Resource not found",
  "status": "error"
}
```
→ Solution: Verify the resource ID/name is correct

---

## Best Practices

### 1. Email Searches
- Use specific terms for better results
- Use `unread_only: true` to filter noise
- Specify folder when you know location

### 2. Calendar Operations
- Always provide start and end times in ISO 8601 format
- Use UTC timezone (Z suffix)
- Check availability before creating events

### 3. Cloud Storage
- Start with `folder_path: "/"` to explore structure
- Use `limit` parameter to avoid large responses
- Cache file IDs for frequently accessed files

### 4. Performance
- Use specific queries instead of broad searches
- Limit results with the `limit` parameter
- Combine operations to reduce API calls

### 5. Error Recovery
- Always check `error` field in responses
- Implement retry logic for rate limits
- Log failed operations for debugging

---

## Combining Tools for Workflows

### Example 1: Email to Calendar
```
1. microsoft_search_emails (find meeting email)
2. microsoft_get_email_details (extract date/time)
3. microsoft_create_calendar_event (add to calendar)
```

### Example 2: Meeting Preparation
```
1. microsoft_get_calendar_events (find meeting)
2. microsoft_search_emails (search for related discussions)
3. microsoft_list_cloud_files (find relevant documents)
```

### Example 3: Email Organization
```
1. microsoft_search_emails (find related emails)
2. microsoft_create_folder (create organization folder)
3. microsoft_manage_email_folders (move emails)
```

---

## Rate Limiting

Microsoft Graph API has rate limits:
- **Standard limits**: 2,000 requests per 10 minutes
- **Large-scale limits**: 4,000 requests per 10 minutes (with throttling)

The server handles rate limiting automatically with:
- Automatic token refresh
- Error messages when limits are exceeded
- Recommendations to retry after delay

---

## Support

For more information:
- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [MCP Protocol Specification](https://modelcontextprotocol.io/)
- Project GitHub Issues

---

**Version**: 1.0.0  
**Last Updated**: 2024
