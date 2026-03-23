# Microsoft 365 MCP Skill for Claude

A **complete, production-ready** Model Context Protocol server that gives Claude direct access to your Microsoft 365 services.

## 🎯 What You Get

- ✅ **10 Powerful Tools**: Email, Calendar, Cloud Storage
- ✅ **Complete Server**: 1,200+ lines of production code
- ✅ **Full Documentation**: Setup guides, API reference, troubleshooting
- ✅ **Easy Setup**: Automated script + Docker support
- ✅ **Secure**: OAuth 2.0, credentials stay on your machine
- ✅ **Professional**: Type-safe, error handling, logging

## 📦 Package Contents

```
microsoft-mcp-skill/
├── SKILL.md                    # Main skill definition
├── FILE_INDEX.md               # Complete file listing
├── scripts/
│   ├── microsoft_mcp.py        # MCP server (1,200+ lines)
│   ├── requirements.txt        # Dependencies
│   └── setup.sh               # Automated setup
├── references/
│   ├── SETUP_GUIDE.md         # Detailed setup (3 phases)
│   ├── API_REFERENCE.md       # Complete API docs
│   ├── CONFIGURATION.md       # Advanced options
│   └── README.md              # Quick start
└── assets/
    ├── .env.template          # Config template
    ├── Dockerfile             # Docker image
    └── docker-compose.yml     # Docker Compose
```

## 🚀 Quick Start (5 Minutes)

### 1. Azure Setup
```bash
# Visit https://portal.azure.com
# Create App Registration → Get Client ID & Secret
# Add permissions (Mail.*, Calendars.*, Files.*)
# Grant admin consent
```

### 2. Install
```bash
# Run automated setup
./scripts/setup.sh

# Or manually install
pip install -r scripts/requirements.txt
cp assets/.env.template .env
# Edit .env with your credentials
```

### 3. Connect to Claude Desktop
```bash
# Edit: ~/.config/Claude/claude_desktop_config.json
# Or: %APPDATA%\Claude\claude_desktop_config.json

"microsoft": {
  "command": "python",
  "args": ["/path/to/scripts/microsoft_mcp.py"],
  "env": {
    "MICROSOFT_CLIENT_ID": "your_id",
    "MICROSOFT_CLIENT_SECRET": "your_secret",
    "MICROSOFT_TENANT_ID": "your_tenant"
  }
}

# Restart Claude Desktop
```

### 4. Start Using!
```
User: "Search my emails for Q1 planning"
Claude: [Searches and shows results]

User: "What's on my calendar this week?"
Claude: [Lists calendar events]

User: "Create a folder in OneDrive called Projects"
Claude: [Creates folder]
```

## 📚 Documentation

- **SKILL.md** - Start here! Overview and quick reference
- **references/SETUP_GUIDE.md** - Detailed 3-phase setup guide
- **references/API_REFERENCE.md** - All 10 tools documented
- **references/CONFIGURATION.md** - Advanced configuration
- **FILE_INDEX.md** - Complete file descriptions

## 🎯 Tools Included

### Email (5 tools)
- Search emails with full-text search
- Send emails with HTML, CC/BCC
- Create and manage drafts
- Get email details and attachments
- Organize folders

### Calendar (3 tools)
- Create events with attendees
- Check availability and conflicts
- List events for date ranges

### Cloud Storage (2 tools)
- Browse OneDrive, SharePoint, Teams
- Create and organize folders

## 🔐 Security

- ✅ Runs locally (not in cloud)
- ✅ OAuth 2.0 (no passwords stored)
- ✅ Credentials never sent to Claude/Anthropic
- ✅ API responses only visible to Claude

## 💡 Usage Examples

### Email Workflow
```
"Find emails about the Q1 budget and extract action items"
→ Searches emails
→ Extracts tasks
→ Shows owners and deadlines
```

### Calendar Integration
```
"Find time to meet with Sarah next week"
→ Checks availability
→ Suggests time slots
→ Creates event
→ Sends notification
```

### Meeting Prep
```
"Prepare me for my 2 PM meeting"
→ Gets event details
→ Searches related emails
→ Lists relevant files
→ Summarizes context
```

## 🎓 Next Steps

1. **Read SKILL.md** - Get an overview
2. **Follow references/SETUP_GUIDE.md** - Complete setup (10 mins)
3. **Run scripts/setup.sh** - Automated installation
4. **Connect to Claude Desktop** - Add to config
5. **Start using!** - Ask Claude to access your services

## 📋 System Requirements

- Python 3.10+
- Microsoft 365 account
- Azure Application (free registration)
- Claude Desktop app

## 🆘 Help

**Setup Issues?**
→ See `references/SETUP_GUIDE.md`

**How do I use a tool?**
→ See `references/API_REFERENCE.md`

**Advanced configuration?**
→ See `references/CONFIGURATION.md`

**General questions?**
→ See `references/README.md`

## 📊 What's Included

| File | Size | Purpose |
|------|------|---------|
| microsoft_mcp.py | 38 KB | Main server code |
| SETUP_GUIDE.md | 15 KB | Installation guide |
| API_REFERENCE.md | 25 KB | Tool documentation |
| Docs & Config | ~20 KB | Templates, guides |
| **Total** | ~102 KB | Complete package |

## ⚡ Features

✅ Production-ready code  
✅ Type-safe with Pydantic validation  
✅ Async/await for performance  
✅ Comprehensive error handling  
✅ OAuth 2.0 authentication  
✅ Automatic token refresh  
✅ Rate limit aware  
✅ Logging & monitoring  
✅ Docker support  
✅ Extensive documentation  

## 🚀 Installation Methods

### Method 1: Automated Setup (Recommended)
```bash
./scripts/setup.sh
# Creates venv, installs deps, sets up .env
```

### Method 2: Manual
```bash
python3 -m venv venv
source venv/bin/activate  # or: venv\Scripts\activate on Windows
pip install -r scripts/requirements.txt
cp assets/.env.template .env
```

### Method 3: Docker
```bash
docker-compose -f assets/docker-compose.yml up -d
```

## 📝 Version Info

- **Name**: microsoft-mcp
- **Version**: 1.0.0
- **Status**: Production Ready
- **Python**: 3.10+
- **Last Updated**: 2024

## 📄 License

MIT - Feel free to use, modify, and distribute

---

## Ready to Get Started?

1. **Extract this folder**
2. **Read `SKILL.md`** for overview
3. **Follow `references/SETUP_GUIDE.md`** for setup
4. **Run `./scripts/setup.sh`** to install
5. **Configure `.env`** with your Azure credentials
6. **Connect to Claude Desktop** (see setup guide)
7. **Start asking Claude to access your Microsoft services!** 🚀

---

**Questions?** Check the comprehensive documentation in the `references/` folder!
