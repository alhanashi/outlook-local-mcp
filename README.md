# outlook-local-mcp

A minimal, dependency-free [MCP](https://modelcontextprotocol.io) server that connects Claude to **Microsoft Outlook for Mac** through AppleScript.

Because it talks to the Outlook app on your Mac rather than to Microsoft's cloud, it works with **on-premises Exchange** mailboxes, where the official Claude for Outlook add-in (which requires Exchange Online / Microsoft Graph) is not supported.

This repository has two parts:

- **MCP server** (root): lets Claude Desktop read your mailbox and create drafts.
- **Outlook add-in** ([`addin/`](addin/)): a task pane inside Outlook that summarizes, classifies and drafts replies for the open message.

## What the MCP server can do

| Tool | Type | Description |
|---|---|---|
| `list_folders` | read | Mail folders with ids and unread counts |
| `list_messages` | read | Recent messages in a folder (defaults to the account Inbox) |
| `search_messages` | read | Search by subject or sender within a date range |
| `get_message` | read | Headers, recipients, attachment names, plain-text body |
| `list_events` | read | Calendar events for the next N days |
| `create_draft` | draft | New email saved as an **unsent** draft and opened for review |
| `create_reply_draft` | draft | Reply / reply-all saved as an **unsent** draft and opened for review |
| `move_message` | organize | Move a message to another folder (e.g. an announcements folder) |

## What it cannot do (by design)

- **Send** mail or invites
- Delete or flag messages, or move them to Deleted Items, Junk, Outbox or Sent
- Change rules, accounts, or settings
- Access the network

Every draft is saved in the account's Drafts folder and opened in Outlook so a human reviews it and presses **Send**.

## Security design

- **Zero dependencies.** Node.js built-ins only; nothing pulled from npm at runtime.
- **No code injection.** All AppleScript is fixed source in `server/index.js`. Inputs are passed as `argv` to `on run argv` via `execFile` (no shell); a sentinel argument prevents values starting with `-` from being parsed as `osascript` options.
- **Input validation.** Ids must be numeric, email addresses are checked against a pattern, and lengths and counts are capped.
- **Untrusted content.** Every response that contains mail content carries a note telling the model not to follow instructions found inside emails (prompt-injection mitigation).
- **No persistence.** The server writes no files and does not log mail content.

The whole server is a single file (~650 lines) and can be audited quickly.

## Requirements

- macOS with **Legacy Outlook** enabled (Outlook menu → *Legacy Outlook*). New Outlook for Mac does not support AppleScript.
- Your mail account added in Legacy Outlook (it keeps its own account list).
- Claude Desktop (it provides the Node.js runtime for extensions), or Node.js 18+.

## Install

### Option A: Claude Desktop extension (.mcpb)

```bash
npx @anthropic-ai/mcpb pack . outlook-local.mcpb
```

Double-click `outlook-local.mcpb` and choose **Install**.

### Option B: manual config

Add to `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook-local": {
      "command": "node",
      "args": ["/absolute/path/to/outlook-local-mcp/server/index.js"]
    }
  }
}
```

On first use, macOS asks to allow Claude to control Microsoft Outlook (System Settings → Privacy & Security → Automation). Allow it.

## Known limitations

- Occurrences of recurring meetings may not all appear in `list_events` (AppleScript returns the series master).
- Search by subject is done by Outlook; search by sender scans messages in the date range, so keep `since_days` reasonable on large folders.
- Large folders can make the first query slow while Outlook evaluates the filter.

## Testing without Outlook

Set `OUTLOOK_MCP_OSASCRIPT` to a mock executable to exercise the JSON-RPC layer on any OS.

## License

MIT
