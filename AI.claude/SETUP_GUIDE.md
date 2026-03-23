# Microsoft MCP Server - Complete Setup Guide

## Overview

This is a production-ready MCP (Model Context Protocol) server that provides Claude with full integration to Microsoft services:
- **Outlook** - Email management (read, send, drafts, folders)
- **Calendar** - Event management and availability checking
- **OneDrive** - Cloud file storage
- **SharePoint** - Team site collaboration
- **Teams** - Team communication

## Architecture

```
Your Machine
    │
    ├─ Microsoft MCP Server (python)
    │   ├─ OAuth 2.0 Token Management
    │   └─ Microsoft Graph API Client
    │
    └─ Claude (via MCP connection)
         └─ Uses MCP tools to access your Microsoft services
```

The MCP server handles all authentication securely - your credentials never leave your machine.

---

## Phase 1: Azure Application Setup

### Step 1: Create Azure App Registration

1. **Go to Azure Portal**
   - URL: https://portal.azure.com
   - Sign in with your Microsoft account

2. **Navigate to App Registrations**
   - Search for "App registrations"
   - Click "New registration"

3. **Register Application**
   - Name: "Claude Microsoft MCP"
   - Supported account types: "Accounts in this organizational directory only"
   - Redirect URI: Leave blank for now (we'll use client credentials flow)
   - Click "Register"

4. **Save These Values** (you'll need them later):
   - **Application (client) ID** → `MICROSOFT_CLIENT_ID`
   - **Directory (tenant) ID** → `MICROSOFT_TENANT_ID`

### Step 2: Create Client Secret

1. **In your app registration, go to "Certificates & secrets"**
2. **Click "New client secret"**
3. **Description:** "Claude MCP"
4. **Expires:** Choose based on your security policy
5. **Copy the Value** → `MICROSOFT_CLIENT_SECRET`
   - ⚠️ Save this immediately - you won't be able to see it again!

### Step 3: Configure API Permissions

1. **Go to "API permissions"**
2. **Click "Add a permission"**
3. **Select "Microsoft Graph"**
4. **Choose "Application permissions"** (for daemon/server access)
5. **Add these permissions:**

**Mail (Email):**
- ☑️ Mail.Read
- ☑️ Mail.ReadWrite
- ☑️ Mail.Send

**Calendar:**
- ☑️ Calendars.Read
- ☑️ Calendars.ReadWrite

**Files:**
- ☑️ Files.Read.All
- ☑️ Files.ReadWrite.All

**Sites:**
- ☑️ Sites.Read.All

**Teams:**
- ☑️ Team.ReadBasic.All
- ☑️ User.Read

6. **Click "Grant admin consent"** (requires admin privileges)
   - If you can't grant admin consent, ask your IT admin

### Step 4: Verify Permissions

Your registered app should now have these permissions with "Granted for..." status showing your organization.

---

## Phase 2: Installation

### Prerequisites

- Python 3.10+
- pip (Python package manager)

### Step 1: Clone or Download the Server

```bash
# Create a directory for the MCP server
mkdir ~/microsoft-mcp
cd ~/microsoft-mcp

# Copy the microsoft_mcp.py file here
cp /path/to/microsoft_mcp.py .

# Copy requirements.txt
cp /path/to/requirements.txt .
```

### Step 2: Install Dependencies

```bash
# Create virtual environment (recommended)
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

### Step 3: Configure Environment Variables

Create a `.env` file in the microsoft-mcp directory:

```bash
cat > .env << 'EOF'
MICROSOFT_CLIENT_ID=your_client_id_here
MICROSOFT_CLIENT_SECRET=your_client_secret_here
MICROSOFT_TENANT_ID=your_tenant_id_here
EOF
```

Or set environment variables directly:

```bash
export MICROSOFT_CLIENT_ID="your_client_id"
export MICROSOFT_CLIENT_SECRET="your_client_secret"
export MICROSOFT_TENANT_ID="your_tenant_id"
```

### Step 4: Test the Server

```bash
python microsoft_mcp.py --help

# Output should show available tools:
# - microsoft_search_emails
# - microsoft_send_email
# - microsoft_create_email_draft
# - microsoft_get_email_details
# - microsoft_manage_email_folders
# - microsoft_create_calendar_event
# - microsoft_check_availability
# - microsoft_get_calendar_events
# - microsoft_list_cloud_files
# - microsoft_create_folder
```

---

## Phase 3: Connect to Claude

### Option A: Using Claude.ai (Web)

1. **Open Claude at https://claude.ai**
2. **Go to Settings → Connected Apps**
3. **Look for "Add custom MCP server" option**
4. **Fill in MCP Configuration:**
   - **Server Name:** Microsoft MCP
   - **Transport:** stdio
   - **Command:** `python /path/to/microsoft_mcp.py`
   - **Environment Variables:**
     ```
     MICROSOFT_CLIENT_ID=your_id
     MICROSOFT_CLIENT_SECRET=your_secret
     MICROSOFT_TENANT_ID=your_tenant
     ```

5. **Test Connection**
   - Click "Test"
   - You should see the list of available tools

### Option B: Using Claude Desktop App

1. **Open Claude Desktop app**
2. **Settings → Developer → Edit Config**
3. **Add to the `mcpServers` section:**

```json
{
  "mcpServers": {
    "microsoft": {
      "command": "python",
      "args": ["/path/to/microsoft_mcp.py"],
      "env": {
        "MICROSOFT_CLIENT_ID": "your_client_id",
        "MICROSOFT_CLIENT_SECRET": "your_client_secret",
        "MICROSOFT_TENANT_ID": "your_tenant_id"
      }
    }
  }
}
```

4. **Save and restart Claude**

### Option C: Using Streamable HTTP Transport (Remote)

For remote deployments:

```bash
# Start server in HTTP mode
python microsoft_mcp.py --transport streamable_http --port 8000

# Then configure Claude to connect to:
# http://localhost:8000
```

---

## Available Tools

### Email Tools

#### 1. `microsoft_search_emails`
Search through emails with full-text search
```
Query: "budget report"
Folder: inbox
Limit: 20
Unread Only: false
```

#### 2. `microsoft_send_email`
Send emails with HTML formatting, CC/BCC
```
To: ["recipient@example.com"]
CC: ["cc@example.com"]
Subject: "Project Update"
Body: "<p>Here's the update...</p>"
Importance: high
```

#### 3. `microsoft_create_email_draft`
Create drafts for later editing
```
Subject: "Meeting Notes"
Body: "Key points from today's meeting..."
```

#### 4. `microsoft_get_email_details`
Get full email content including attachments
```
Message ID: "AAMkADE4..."
Include Attachments: true
```

#### 5. `microsoft_manage_email_folders`
Manage email folders
```
Action: list | create | delete | rename
Folder Name: "Important"
New Name: "Critical" (for rename)
```

### Calendar Tools

#### 6. `microsoft_create_calendar_event`
Create calendar events with attendees
```
Subject: "Team Standup"
Start: "2024-01-15T10:00:00Z"
End: "2024-01-15T10:30:00Z"
Attendees: ["alice@example.com", "bob@example.com"]
```

#### 7. `microsoft_check_availability`
Check your availability for a time slot
```
Start: "2024-01-15T14:00:00Z"
End: "2024-01-15T15:00:00Z"
Attendees: ["alice@example.com"] (optional)
```

#### 8. `microsoft_get_calendar_events`
List calendar events for a date range
```
Start Date: "2024-01-15"
End Date: "2024-01-22"
Limit: 50
Include Details: true
```

### Cloud Storage Tools

#### 9. `microsoft_list_cloud_files`
List files in OneDrive, SharePoint, or Teams
```
Location: onedrive | sharepoint | teams
Folder Path: "/Projects"
Limit: 20
```

#### 10. `microsoft_create_folder`
Create new folders in cloud storage
```
Folder Name: "Q1 Planning"
Parent Path: "/Projects"
Location: onedrive
```

---

## Usage Examples

### Example 1: Email Workflow
```
User: "Search my inbox for emails about budget reports"
Claude: [Uses microsoft_search_emails with query="budget report"]

User: "Show me the full details of the top result"
Claude: [Uses microsoft_get_email_details with the message ID]

User: "Send a reply acknowledging receipt"
Claude: [Uses microsoft_send_email to send reply]
```

### Example 2: Calendar with Context
```
User: "Find a 1-hour slot next week to meet with Alice"
Claude: [Uses microsoft_check_availability for multiple time slots]
Claude: [Shows available slots and creates event when approved]

User: "Add this to my calendar and email Alice about the meeting"
Claude: [Uses microsoft_create_calendar_event]
Claude: [Uses microsoft_send_email to notify Alice]
```

### Example 3: Email-Driven Task Management
```
User: "Look through my emails for action items and create a summary"
Claude: [Uses microsoft_search_emails with various queries]
Claude: [Extracts action items from email contents]
Claude: [Provides structured summary with ownership]
```

---

## Troubleshooting

### Connection Issues

**Problem:** "Cannot connect to Microsoft services"
```
Solution:
1. Verify environment variables are set correctly
2. Check that the Azure app has the correct permissions
3. Ensure "Grant admin consent" was clicked for all permissions
4. Try refreshing the token by restarting the server
```

**Problem:** "Permission denied" errors
```
Solution:
1. Go to Azure Portal → Your App → API Permissions
2. Verify all required permissions are listed
3. Click "Grant admin consent" again
4. Wait a few minutes for permissions to propagate
```

**Problem:** "Invalid client credentials"
```
Solution:
1. Verify MICROSOFT_CLIENT_SECRET hasn't expired
2. Create a new client secret if it has
3. Check that CLIENT_ID and TENANT_ID are correct
4. Ensure no extra spaces in environment variables
```

### Tool Issues

**Problem:** Email search returns no results
```
Solution:
1. Try broader search terms
2. Check the specified folder exists
3. Verify you have mail.read permission
4. Try searching for sender address specifically
```

**Problem:** Can't send emails
```
Solution:
1. Verify "Mail.Send" permission is granted
2. Ensure recipient email addresses are valid
3. Check email body doesn't exceed limits
4. Verify "Grant admin consent" was completed
```

**Problem:** Calendar events not showing
```
Solution:
1. Verify date range includes current/future events
2. Check calendar sharing settings
3. Ensure "Calendars.Read" permission is granted
4. Try expanding date range
```

---

## Security Best Practices

### 1. Environment Variables
- ✅ Use `.env` file (add to `.gitignore`)
- ✅ Use environment variables from secure vaults
- ❌ Don't hardcode credentials in code
- ❌ Don't commit `.env` to git

### 2. Client Secret Management
- ✅ Store securely (e.g., 1Password, Vault, AWS Secrets Manager)
- ✅ Rotate periodically
- ❌ Share with others
- ❌ Log or display in output

### 3. API Usage
- ✅ Monitor API usage for anomalies
- ✅ Use principle of least privilege
- ✅ Regularly audit connected apps
- ❌ Grant more permissions than needed

### 4. Server Deployment
- ✅ Run behind HTTPS
- ✅ Use firewall rules
- ✅ Monitor server logs
- ❌ Expose to public internet without auth

---

## Advanced Configuration

### Increasing Rate Limits
Microsoft Graph has rate limits. For production use:

```python
# In microsoft_mcp.py, modify the HTTP client
async with httpx.AsyncClient(
    timeout=30.0,
    limits=httpx.Limits(max_keepalive_connections=5)
) as client:
    # Your requests here
```

### Caching Tokens
The current implementation caches tokens in memory. For production:

```python
# Use a file-based cache or database
import json
import os

class TokenCache:
    def __init__(self, cache_file=".token_cache"):
        self.cache_file = cache_file
    
    def get(self):
        if os.path.exists(self.cache_file):
            with open(self.cache_file, 'r') as f:
                return json.load(f)
        return None
    
    def save(self, token_data):
        with open(self.cache_file, 'w') as f:
            json.dump(token_data, f)
```

### Multi-Tenant Support
For organizations with multiple tenants:

```python
class MultiTenantGraphAPIClient:
    def __init__(self, configs: Dict[str, MicrosoftAuthConfig]):
        self.clients = {
            name: GraphAPIClient(config)
            for name, config in configs.items()
        }
```

---

## Monitoring & Logging

### Enable Detailed Logging

```python
import logging

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# In your tools:
logger.info(f"Searching emails with query: {params.query}")
logger.debug(f"API response: {response}")
```

### Monitor API Usage

```python
# Track API calls
api_calls = {
    "search_emails": 0,
    "send_email": 0,
    # ... etc
}

@mcp.tool(name="microsoft_search_emails")
async def search_emails(params, ctx):
    api_calls["search_emails"] += 1
    # ... rest of function
```

---

## Getting Help

### Common Questions

**Q: Can I use this with my personal Microsoft account?**
A: Yes, but you need to create an Azure application first. Personal accounts require additional setup.

**Q: Is my data safe?**
A: Yes. The server runs locally, and credentials are never sent to Claude or Anthropic - only the API responses.

**Q: Can I modify the server to add more tools?**
A: Absolutely! The server is open-source and extensible. See the [MCP Protocol](https://modelcontextprotocol.io) for details.

**Q: What if I don't have admin privileges to grant consent?**
A: Contact your IT administrator. They can grant admin consent for the application.

---

## Next Steps

1. ✅ Complete Azure setup (Phase 1)
2. ✅ Install the server locally (Phase 2)
3. ✅ Connect to Claude (Phase 3)
4. 🚀 Start using the tools!

Example first command:
```
"Search my emails for 'Q1 planning' and show me what we discussed"
```

---

## Version & Updates

- **Current Version:** 1.0.0
- **Last Updated:** 2024
- **Python Version:** 3.10+

Check for updates regularly for new features and security patches.

---

Happy organizing! 🚀
