---
name: microsoft-mcp
description: |
  Complete MCP server for Microsoft 365 integration with Claude. Enables email management (search, send, draft), calendar operations (create events, check availability), and cloud file access (OneDrive, SharePoint, Teams). Use this skill whenever a user needs to access their Outlook emails, manage calendar events, check meeting availability, search email attachments, organize files in cloud storage, or link email discussions to calendar events. Include setup guidance for Azure authentication.
compatibility:
  - Python 3.10+
  - Microsoft 365 (Office 365) account
  - Azure Application Registration (free)
---

# Microsoft 365 MCP Server for Claude

A production-ready Model Context Protocol server that gives Claude direct access to your Microsoft 365 services: **Outlook, Calendar, OneDrive, SharePoint, and Teams**.

## When to Use This Skill

Use this skill whenever the user mentions:
- **Email management**: "Search my emails", "Send an email", "Create a draft"
- **Calendar access**: "What's on my calendar?", "Find a meeting time", "Create an event"
- **Cloud storage**: "List my OneDrive files", "Create a folder", "Find documents"
- **Cross-service workflows**: "Link this email to calendar", "Show meeting context", "Extract action items from emails"
- **Meeting preparation**: "Prepare for my 2 PM meeting", "Get the background on this discussion"

## What This Skill Enables

### 📧 Email Management (5 Tools)
- Full-text email search with folder and unread filtering
- Send emails with HTML formatting, CC/BCC, and importance levels
- Create and manage email drafts
- Retrieve full email content with attachments
- Organize emails into custom folders

### 📅 Calendar Operations (3 Tools)
- Create calendar events with attendees and reminders
- Check availability for time slots and conflicts
- List events for specific date ranges
- Get full event details with organizer information

### ☁️ Cloud Storage (2 Tools)
- Browse files on OneDrive, SharePoint, and Teams
- Create and organize folders
- Get file metadata (size, modified date, sharing URLs)

### 🤝 Contextual Intelligence
- Extract action items from email conversations
- Link emails to calendar events
- Track meeting decisions from email threads
- Multi-step workflows combining email and calendar data

---

## Quick Start (5 Minutes)

### Step 1: Azure Application Setup

1. Go to **Azure Portal**: https://portal.azure.com
2. Create new **App Registration** (name: "Claude Microsoft MCP")
3. Get these values:
   - **Application (client) ID**
   - **Directory (tenant) ID**
4. Create **Client Secret** (Certificates & Secrets)
5. **Add API Permissions**:
   - Mail.Read, Mail.ReadWrite, Mail.Send
   - Calendars.Read, Calendars.ReadWrite
   - Files.Read.All, Files.ReadWrite.All
   - Sites.Read.All, Team.ReadBasic.All
6. **Grant admin consent**

### Step 2: Install the Server

```bash
# Navigate to the skill files
cd microsoft-mcp-skill

# Install dependencies
pip install -r scripts/requirements.txt

# Configure credentials
cp assets/.env.template .env
# Edit .env with your Azure credentials
```

### Step 3: Connect to Claude Desktop

**Mac/Linux:**
```bash
# Edit Claude Desktop config
open ~/.config/Claude/claude_desktop_config.json

# Add to the "mcpServers" section:
"microsoft": {
  "command": "python",
  "args": ["/path/to/scripts/microsoft_mcp.py"],
  "env": {
    "MICROSOFT_CLIENT_ID": "your_id",
    "MICROSOFT_CLIENT_SECRET": "your_secret",
    "MICROSOFT_TENANT_ID": "your_tenant"
  }
}
```

**Windows:**
```json
Edit: %APPDATA%\Claude\claude_desktop_config.json

"microsoft": {
  "command": "python",
  "args": ["C:\\path\\to\\scripts\\microsoft_mcp.py"],
  "env": {
    "MICROSOFT_CLIENT_ID": "your_id",
    "MICROSOFT_CLIENT_SECRET": "your_secret",
    "MICROSOFT_TENANT_ID": "your_tenant"
  }
}
```

4. **Restart Claude Desktop** - tools will now be available!

---

## Available Tools (10 Total)

### Email Tools (5)
- **microsoft_search_emails** - Full-text search with folder & unread filtering
- **microsoft_send_email** - Send with HTML, CC/BCC, importance levels
- **microsoft_create_email_draft** - Create editable drafts
- **microsoft_get_email_details** - Get full content & attachments
- **microsoft_manage_email_folders** - List, create, delete, rename folders

### Calendar Tools (3)
- **microsoft_create_calendar_event** - Create events with attendees
- **microsoft_check_availability** - Check availability & conflicts
- **microsoft_get_calendar_events** - List events for date ranges

### Cloud Storage Tools (2)
- **microsoft_list_cloud_files** - Browse OneDrive, SharePoint, Teams
- **microsoft_create_folder** - Create folders in cloud storage

---

## Usage Examples

### Email Search & Action Items
```
User: "Find emails about Q1 planning and extract action items"

Claude will:
1. Search emails for "Q1 planning"
2. Get full email details
3. Extract action items, owners, deadlines
4. Present organized summary
```

### Schedule Meeting
```
User: "Find a time to meet with Alice next week and send her an invite"

Claude will:
1. Check your availability next week
2. Suggest available time slots
3. Create calendar event
4. Send email notification
```

### Meeting Preparation
```
User: "Prepare me for the design team meeting at 2 PM"

Claude will:
1. Get calendar event details
2. Search related emails
3. List relevant files from SharePoint
4. Summarize context and discussion points
```

---

## Reference Documentation

For detailed information, see the reference files:

- **SETUP_GUIDE.md** - Comprehensive 3-phase setup guide
  - Azure application registration steps (with screenshots)
  - Local installation instructions
  - Claude Desktop configuration
  - Troubleshooting for common issues

- **API_REFERENCE.md** - Complete tool documentation
  - All 10 tools with detailed parameters
  - Response formats and examples
  - Error codes and solutions
  - Rate limiting information
  - Best practices

---

## Security & Privacy

Your data is completely safe:
- ✅ Server runs locally (not in cloud)
- ✅ Uses OAuth 2.0 (no password storage)
- ✅ Credentials never sent to Claude/Anthropic
- ✅ Only API responses visible to Claude

Best practices:
- Store `.env` file securely (add to `.gitignore`)
- Never commit credentials to version control
- Rotate client secrets periodically
- Monitor API usage
- Review permissions annually

---

## What's Included

- **scripts/microsoft_mcp.py** - Full MCP server (1,200+ lines)
- **scripts/requirements.txt** - Python dependencies
- **scripts/setup.sh** - Automated setup script
- **assets/.env.template** - Configuration template
- **assets/Dockerfile** - Docker containerization
- **assets/docker-compose.yml** - Docker Compose setup
- **references/SETUP_GUIDE.md** - Detailed setup instructions
- **references/API_REFERENCE.md** - Complete API reference
- **references/CONFIGURATION.md** - Advanced configuration

---

## Next Steps

1. Read **references/SETUP_GUIDE.md** for Azure setup (10 minutes)
2. Run `pip install -r scripts/requirements.txt`
3. Configure `.env` with your credentials
4. Connect to Claude Desktop
5. Start using! Example: "Search my emails for important clients"

---

## Troubleshooting Quick Links

See **references/SETUP_GUIDE.md** for:
- Permission denied errors
- Invalid credentials
- Email search returning nothing
- Calendar events not showing
- Connection issues

---

## Version & Support

- **Version**: 1.0.0
- **Status**: Production Ready
- **Python**: 3.10+
- **Last Updated**: 2024

For issues or questions, refer to the comprehensive documentation in the references folder.

---

**Enjoy seamless Microsoft 365 integration with Claude!** 🚀
