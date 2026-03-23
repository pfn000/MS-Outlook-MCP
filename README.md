# Microsoft MCP Server

A production-ready Model Context Protocol server that enables Claude to integrate with your Microsoft 365 services - **Outlook, Calendar, OneDrive, SharePoint, and Teams**.

## 🎯 What This Does

This MCP server gives Claude the ability to:

### 📧 Email (Outlook)
- Search emails with full-text search
- Read email details (subject, body, attachments, recipients)
- Send emails with HTML formatting, CC/BCC
- Create and manage email drafts
- Organize emails into folders
- Get email metadata (sender, date, importance)

### 📅 Calendar
- View your calendar events
- Create new events with attendees
- Check your availability for time slots
- See scheduled meetings and conflicts
- Set reminders for events

### ☁️ Cloud Storage
- List files on OneDrive
- Explore SharePoint sites
- Access Teams file storage
- Create folders and organize files
- Get file metadata (size, modified date, URLs)

### 🤝 Contextual Intelligence
The server understands context:
- Extract action items from emails
- Link emails to calendar events
- Track meeting decisions
- Manage multi-step email workflows
- Cross-reference calendar with emails

## 🚀 Quick Start (5 Minutes)

### 1. Set Up Azure Application

```bash
# Go to https://portal.azure.com
# 1. Create new App Registration ("Claude Microsoft MCP")
# 2. Copy these values:
#    - Application (client) ID
#    - Directory (tenant) ID
# 3. Create Client Secret (Certificates & Secrets)
# 4. Add API Permissions:
#    ☑️ Mail.Read, Mail.ReadWrite, Mail.Send
#    ☑️ Calendars.Read, Calendars.ReadWrite
#    ☑️ Files.Read.All, Files.ReadWrite.All
#    ☑️ Sites.Read.All
# 5. Click "Grant admin consent"
```

### 2. Install the Server

```bash
# Clone or download files
mkdir ~/microsoft-mcp
cd ~/microsoft-mcp

# Copy files here:
# - microsoft_mcp.py
# - requirements.txt
# - .env.template

# Install dependencies
pip install -r requirements.txt

# Set up configuration
cp .env.template .env
# Edit .env with your Azure credentials
```

### 3. Connect to Claude

**Claude.ai (Web):**
- Settings → Connected Apps
- Add custom MCP server
- Command: `python /path/to/microsoft_mcp.py`
- Environment variables: Add your Azure credentials

**Claude Desktop:**
- Settings → Developer → Edit Config
- Add to `mcpServers`:
```json
{
  "microsoft": {
    "command": "python",
    "args": ["/path/to/microsoft_mcp.py"],
    "env": {
      "MICROSOFT_CLIENT_ID": "your_id",
      "MICROSOFT_CLIENT_SECRET": "your_secret",
      "MICROSOFT_TENANT_ID": "your_tenant"
    }
  }
}
```

### 4. Start Using!

```
User: "Search my inbox for emails about the product roadmap"
Claude: [Searches and shows relevant emails]

User: "Find me a 1-hour slot next week to meet with Sarah"
Claude: [Checks calendar, suggests available times, creates event]

User: "What action items came out of yesterday's meeting?"
Claude: [Searches emails, extracts tasks, shows owners]
```

## 📋 Available Tools

### Email Management (5 tools)
| Tool | Purpose |
|------|---------|
| `microsoft_search_emails` | Full-text email search with folder/unread filters |
| `microsoft_send_email` | Send emails with HTML, CC/BCC, importance levels |
| `microsoft_create_email_draft` | Create draft emails for later editing |
| `microsoft_get_email_details` | Get full email content and attachments |
| `microsoft_manage_email_folders` | Create, delete, rename, list folders |

### Calendar Management (3 tools)
| Tool | Purpose |
|------|---------|
| `microsoft_create_calendar_event` | Create events with attendees and reminders |
| `microsoft_check_availability` | Check time slot availability |
| `microsoft_get_calendar_events` | List events for date ranges |

### Cloud Storage (2 tools)
| Tool | Purpose |
|------|---------|
| `microsoft_list_cloud_files` | Browse OneDrive, SharePoint, Teams files |
| `microsoft_create_folder` | Create organized folder structures |

## 🏗️ Architecture

```
Your Computer
    │
    ├─ Python MCP Server
    │   ├─ OAuth Token Manager
    │   ├─ Microsoft Graph Client
    │   └─ 10 Domain Tools
    │
    └─ Claude (local or claude.ai)
         └─ Uses tools to access your Microsoft services
```

**Key Points:**
- ✅ Runs locally - your credentials stay on your machine
- ✅ OAuth 2.0 - secure authentication
- ✅ No data sent to Anthropic - only API responses
- ✅ Automatic token refresh
- ✅ Rate-limited error handling

## 🔐 Security

### Your Data is Safe
- **Local Execution**: MCP server runs on your computer
- **Encrypted Credentials**: Uses OAuth 2.0, not passwords
- **No Data Sharing**: Claude only sees API responses you need
- **Audit Trail**: All API calls can be logged

### Best Practices
```bash
# ✅ Use .env file (add to .gitignore)
cat > .env << EOF
MICROSOFT_CLIENT_ID=your_id
MICROSOFT_CLIENT_SECRET=your_secret
MICROSOFT_TENANT_ID=your_tenant
EOF

# ✅ Never commit credentials
echo ".env" >> .gitignore

# ✅ Keep client secret secure
# Don't share it, don't log it, don't expose it

# ❌ Don't hardcode credentials
# ❌ Don't commit to git
# ❌ Don't share with anyone
```

## 📚 Documentation

- **[SETUP_GUIDE.md](./SETUP_GUIDE.md)** - Detailed setup instructions
- **[.env.template](./.env.template)** - Configuration template
- **[microsoft_mcp.py](./microsoft_mcp.py)** - Full server implementation

## 🧪 Testing

```bash
# Verify server works
python microsoft_mcp.py --help

# Test with sample curl command (after server starts)
curl -X POST http://localhost:8000 \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc": "2.0", "method": "tools/list", "id": 1}'
```

## ⚙️ Advanced Usage

### Environment Variables
```bash
# Required
MICROSOFT_CLIENT_ID=...
MICROSOFT_CLIENT_SECRET=...
MICROSOFT_TENANT_ID=...

# Optional
MCP_PORT=8000          # HTTP server port
MCP_HOST=localhost     # HTTP server host
DEBUG=false            # Enable verbose logging
```

### Custom Configurations
```bash
# Run with custom port
python microsoft_mcp.py --port 9000

# Run with verbose logging
DEBUG=true python microsoft_mcp.py

# Run in HTTP mode (instead of stdio)
python microsoft_mcp.py --transport streamable_http
```

### Extending the Server
The server is extensible. Add new tools by:

```python
from pydantic import BaseModel, Field

class MyInput(BaseModel):
    param: str = Field(..., description="Parameter description")

@mcp.tool(name="my_tool")
async def my_tool(params: MyInput, ctx) -> str:
    client = ctx.request_context.lifespan_state["client"]
    # Use client to call Microsoft Graph API
    result = await client.get("/me/endpoint")
    return json.dumps(result)
```

## 🐛 Troubleshooting

### "Permission denied" errors
→ Go to Azure → App → API Permissions → Click "Grant admin consent"

### "Invalid credentials"
→ Check environment variables are set correctly and secret hasn't expired

### Email search returns nothing
→ Try broader search terms or different folder names

### Calendar events not showing
→ Expand date range, verify calendar sharing settings

See **[SETUP_GUIDE.md](./SETUP_GUIDE.md)** for more troubleshooting.

## 📞 Support

### Common Issues
1. **Azure Setup**: Review Azure Portal steps in SETUP_GUIDE.md
2. **Connection**: Verify environment variables and permissions
3. **API Errors**: Check Microsoft Graph documentation
4. **Python Issues**: Ensure Python 3.10+ and dependencies installed

### Resources
- [Microsoft Graph API Docs](https://docs.microsoft.com/en-us/graph/)
- [Azure App Registration Guide](https://docs.microsoft.com/en-us/azure/active-directory/develop/)
- [MCP Protocol Specification](https://modelcontextprotocol.io/)

## 📝 License

MIT License - Feel free to modify and distribute

## 🎉 Features

✅ **10 Powerful Tools**  
✅ **Production-Ready Code**  
✅ **Comprehensive Error Handling**  
✅ **Type-Safe with Pydantic**  
✅ **Async/Await for Performance**  
✅ **Detailed Documentation**  
✅ **Security Best Practices**  
✅ **Easy Setup & Configuration**  
✅ **Extensible Architecture**  

## 🚀 Next Steps

1. Follow the [5-minute quick start](#-quick-start-5-minutes)
2. Read [SETUP_GUIDE.md](./SETUP_GUIDE.md) for detailed instructions
3. Connect to Claude and start using the tools
4. Explore contextual workflows (emails → calendar → tasks)

---

## Example Workflows

### Workflow 1: Email-to-Calendar
```
User: "Find emails about the Q1 planning meeting and add it to my calendar"
Claude:
  1. Searches emails for "Q1 planning"
  2. Extracts meeting date/time
  3. Identifies attendees
  4. Creates calendar event
  5. Shows confirmation
```

### Workflow 2: Email-based Action Items
```
User: "Show me action items from emails sent by my manager this week"
Claude:
  1. Searches emails from manager
  2. Analyzes content for action items
  3. Extracts deadlines from calendar
  4. Creates prioritized list
```

### Workflow 3: Meeting Preparation
```
User: "Prepare me for my meeting with the design team at 2 PM"
Claude:
  1. Gets calendar event details
  2. Searches relevant emails
  3. Checks attendee availability
  4. Summarizes background context
  5. Lists agenda items
```

---

**Enjoy seamless Microsoft 365 integration with Claude!** 🎉
