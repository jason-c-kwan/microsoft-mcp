# Microsoft MCP Server - Qwen Code Context

## Project Overview
This is an MCP (Model Context Protocol) server that provides AI assistants with comprehensive access to Microsoft Graph APIs, enabling integration with Outlook email, Calendar, OneDrive, and Contacts.

## Key Architecture Components

### Core Files
- **`src/microsoft_mcp/server.py`**: Main entry point that starts the MCP server
- **`src/microsoft_mcp/tools.py`**: Contains all MCP tool definitions and business logic (30+ tools)
- **`src/microsoft_mcp/auth.py`**: Handles Microsoft authentication using MSAL
- **`src/microsoft_mcp/graph.py`**: Low-level Microsoft Graph API client
- **`authenticate.py`**: Standalone authentication helper script

### Authentication System
- Uses Microsoft Authentication Library (MSAL) for OAuth2 device flow authentication
- Supports multi-account authentication (personal, work, school)
- Token cache stored in `~/.microsoft_mcp_token_cache.json`
- Requires Azure App Registration with specific permissions

### Required Azure Permissions
- Mail.ReadWrite
- Calendars.ReadWrite  
- Files.ReadWrite
- Contacts.Read
- People.Read
- User.Read

## Available Tools (30+ tools)

### Email Management
- `list_emails` - List emails with optional body content
- `get_email` - Get specific email with attachments
- `create_email_draft` - Create drafts with attachment support
- `send_email` - Send emails with CC/BCC and attachments
- `reply_to_email` / `reply_all_email` - Reply maintaining thread context
- `update_email` - Mark as read/unread
- `move_email` - Move between folders
- `delete_email` - Delete emails
- `get_attachment` - Get attachment content
- `semantic_search_emails` - **PREFERRED** - Natural language email search with relevance ranking
- `search_emails_advanced` - Structured search with filters (dates, senders, folders)
- `search_emails` - Basic keyword search (fallback)

### Calendar Management
- `list_events` - List calendar events
- `get_event` - Get event details
- `create_event` - Create events with attendees, location, categories
- `update_event` - Modify existing events
- `delete_event` - Cancel events
- `respond_event` - Accept/decline/tentative responses
- `check_availability` - Check free/busy times
- `list_events_by_date_range` - List events within specific date range with category filtering
- `semantic_search_calendar` - **PREFERRED** - Natural language calendar search with relevance ranking
- `search_events` - Basic keyword search (fallback)

### Contact Management
- `list_contacts` - List all contacts
- `get_contact` - Get contact details
- `create_contact` - Create new contacts
- `update_contact` - Update contact info
- `delete_contact` - Delete contacts
- `search_contacts` - Search contacts

### File Management (OneDrive)
- `list_files` - Browse OneDrive files/folders
- `get_file` - Download file content
- `create_file` - Upload files
- `update_file` - Update existing files
- `delete_file` - Delete files/folders
- `search_files` - Search OneDrive files

### Category Management (Calendar)
- `list_outlook_categories` - List all outlook categories
- `create_outlook_category` - Create new categories with colors
- `update_outlook_category_color` - Update category colors
- `delete_outlook_category` - Delete categories
- `get_outlook_category_by_name` - Find category by display name
- `list_available_colors` - List all available colors for categories

### Search Tools
- `semantic_unified_search` - **PREFERRED** - Natural language search across emails, events, and files
- `unified_search` - Basic keyword search across resources (fallback)
- `semantic_search_emails` - Natural language email search
- `semantic_search_calendar` - Natural language calendar search
- `search_emails_advanced` - Structured email search with filters

### Utility Tools
- `list_accounts` - Show authenticated accounts
- `authenticate_account` - Start new account authentication
- `complete_authentication` - Complete auth flow
- `debug_token_info` - Debug function to inspect token scopes and claims
- `test_mail_permissions` - Test specific mail permissions to diagnose 403 errors

### Search Tool Hierarchy
**ALWAYS use semantic search tools first for better results:**
1. **For targeted searches**: `semantic_search_emails` / `semantic_search_calendar` - Natural language queries for specific resource types
2. **For cross-resource searches**: `semantic_unified_search` - Natural language queries across emails, events, and files
3. **For structured filtering**: `search_emails_advanced` - Precise filters (dates, senders, folders)
4. **Fallback only**: `search_emails` / `search_events` / `unified_search` - Basic keyword matching

## Color System for Calendar Categories
Supports 25 color presets mapped to natural language:
- Standard: red, orange, brown, yellow, green, teal, olive, blue, purple, cranberry, steel, darksteel, gray, darkgray, black
- Dark variants: darkred, darkorange, darkbrown, darkyellow, darkgreen, darkteal, darkolive, darkblue, darkpurple, darkcranberry
- Space aliases: "dark red", "dark orange", etc.

## Development Commands
```bash
# Install dependencies
uv sync

# Run tests
uv run pytest tests/ -v

# Type checking
uv run pyright

# Format code
uvx ruff format .

# Lint
uvx ruff check --fix --unsafe-fixes .

# Run MCP server locally
uv run microsoft-mcp
```

## Environment Variables
- `MICROSOFT_MCP_CLIENT_ID` - Required Azure App ID
- `MICROSOFT_MCP_TENANT_ID` - Optional (defaults to "common")

## Multi-Account Support
All tools require `account_id` as first parameter. Use `list_accounts()` to get available account IDs.

## Common Usage Patterns
1. **Authentication**: Always get account_id first with `list_accounts()`
2. **Email workflows**: List → Get → Reply/Forward/Move
3. **Calendar workflows**: Check availability → Create event → Send invitations
4. **File workflows**: List → Get/Upload → Share/Organize
5. **Search**: Use unified_search for cross-service queries

## Error Handling
- Graph API errors are caught and re-raised with context
- Authentication failures prompt re-authentication
- Rate limiting handled with exponential backoff
- Token refresh handled automatically

## Security Notes
- Tokens cached locally, never logged
- All API calls use HTTPS
- Permissions follow principle of least privilege
- Support for enterprise/organizational constraints

## Testing
- Integration tests in `tests/test_integration.py`
- Requires real Microsoft account for full testing
- Mock testing for unit tests (when available)

## Current Branch Status
- Branch: `autogen_version` (created for this session)
- Based on recent work with advanced email search functionality and semantic search improvements

## Repository Structure
```
.
├── authenticate.py                 # Authentication helper script
├── pyproject.toml                  # Project configuration
├── README.md                       # Main documentation
├── src/
│   └── microsoft_mcp/
│       ├── __init__.py
│       ├── auth.py                 # Authentication logic
│       ├── graph.py                # Microsoft Graph API client
│       ├── server.py               # MCP server entry point
│       └── tools.py                # All MCP tools (main logic)
└── tests/
    └── test_integration.py         # Integration tests
```

## Dependencies
- Python 3.12+
- fastmcp>=2.8.0
- httpx>=0.28.1
- msal>=1.32.3
- python-dotenv>=1.1.0
- mcp>=1.9.3 (dev)
- pytest>=8.4.0 (dev)
- pytest-asyncio>=1.0.0 (dev)
