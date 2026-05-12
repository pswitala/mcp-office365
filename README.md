# mcp-claude-office365

MCP server for Microsoft Office 365 -- inbox and calendar integration for Claude.

## Features

- **Email**: list, search, send, reply, create drafts, move messages, list folders
- **Calendar**: list, create, update, delete events, find meeting times, respond to invitations, list calendars
- **Multi-account**: configure and switch between multiple O365 mailboxes
- **Resources**: read-only `o365://inbox`, `o365://calendar-today`, `o365://calendar-week` endpoints

## Prerequisites

An Azure AD application registration with the following:

1. Register a new app at [portal.azure.com](https://portal.azure.com)
2. Supported account types: **Accounts in any organizational directory (multi-tenant) and personal Microsoft accounts**
3. Authentication: Add platform **Mobile and desktop** and enable **Allow public client flows**
4. API permissions (Delegated):
   - `Mail.Read`
   - `Mail.Send`
   - `Mail.ReadWrite`
   - `Calendars.ReadWrite`
   - `User.Read`
5. Note the **Application (client) ID** -- you'll need it below

## Installation

```bash
npm install
npm run build
```

## Configuration

### 1. Set your Azure AD client ID

Use the `configure` tool in your MCP client:

```json
{ "name": "configure", "arguments": { "clientId": "your-azure-ad-client-id-here" } }
```

The client ID is persisted to `~/.mcp-o365/config.json` so you only need to do this once.

### 2. Authenticate a mailbox

Use the `login` tool with a display alias:

```json
{ "name": "login", "arguments": { "alias": "work" } }
```

This triggers the Microsoft device code flow:
- A URL and code are printed to the console
- Open the URL in a browser and enter the code
- Sign in and grant permissions

The account is saved to `~/.mcp-o365/accounts.json` and can be re-used across sessions (until the access token expires).

### 3. Verify status

```json
{ "name": "status" }
```

Shows configured client ID, list of aliases, and token validity.

## Tools

### Mail

| Tool | Description |
|------|-------------|
| `list_messages` | List emails from a folder (default: inbox). Supports search and pagination. |
| `get_message` | Get full content of an email by ID. |
| `search_messages` | Search emails using keyword query. |
| `send_message` | Send an email (text or HTML body). Supports CC. |
| `reply_message` | Reply to an existing message. |
| `create_draft` | Create a draft email. |
| `move_message` | Move a message to another folder. |
| `list_folders` | List mail folders in a mailbox. |

### Calendar

| Tool | Description |
|------|-------------|
| `list_events` | List calendar events in a date range. |
| `get_event` | Get full details of an event by ID. |
| `create_event` | Create a new calendar event with optional attendees and location. |
| `update_event` | Partial update of an existing event. |
| `delete_event` | Cancel/delete an event. |
| `find_meeting_times` | Find available meeting times for a set of attendees. |
| `list_calendars` | List all calendars for a mailbox. |
| `respond_to_event` | Accept, tentatively accept, or decline a calendar invitation. |

### Auth

| Tool | Description |
|------|-------------|
| `configure` | Set the Azure AD client ID. |
| `login` | Authenticate a mailbox (device code flow). |
| `check_login` | Verify a saved alias has a valid token. |
| `list_accounts` | List all configured mailbox aliases. |
| `remove_account` | Remove a mailbox alias. |
| `status` | Show current configuration and token status. |

## Resources

| URI | Description |
|-----|-------------|
| `o365://inbox` | Last 10 unread messages across all configured mailboxes |
| `o365://calendar-today` | Today's calendar events |
| `o365://calendar-week` | Calendar events for the next 7 days |

## Usage with MCP Clients

### Claude Code (CLI)

Add to your MCP configuration (e.g. `.mcp.json` or equivalent):

```json
{
  "mcpServers": {
    "office365": {
      "command": "node",
      "args": ["/path/to/mcp-office365/dist/index.js"],
      "env": {
        "MCP_O365_CLIENT_ID": "your-client-id"
      }
    }
  }
}
```

Alternatively, skip the env var by running the `configure` tool once after connecting.

### Development mode

```bash
npm run dev   # watches src/ and restarts on changes
npm run start # runs compiled dist/index.js
```

## Configuration Files

All persistent data is stored under `~/.mcp-o365/`:

| File | Contents |
|------|----------|
| `config.json` | Azure AD client ID |
| `accounts.json` | Authenticated mailbox aliases (account IDs, usernames) |

> **Note:** Tokens are cached by MSAL in the same directory. Access tokens expire after ~1 hour. When expired, re-authenticate using the `login` tool.

## Troubleshooting

- **"No Office 365 client ID configured"** -- Run the `configure` tool or set `MCP_O365_CLIENT_ID`.
- **"No valid token for alias"** -- Run the `login` tool again to re-authenticate.
- **Permission denied errors** -- Ensure your Azure AD app has all the required API permissions listed above and that an admin has consented if required by your tenant.
