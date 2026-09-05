# Outlook (Windows)

**Automate Microsoft Outlook with Claude** - work with your mail, calendar, folders and attachments through natural language. Requires Windows and a local classic Outlook install.

## What It Does

- **Mail** - read the message you have open, search and list messages, draft, reply, reply-all, forward, send, move, delete, and set read state, categories, subject, body or recipients
- **Calendar** - list and read appointments, create, update and delete them
- **Contacts** - list and read contacts, create, update and delete them
- **Tasks** - list and read tasks, create, update and delete them
- **Rules** - list the inbox rules that decide what happens to mail before you read it, and create, change, switch off or remove them
- **Folders** - list default folders, walk child folders, resolve a folder path, list items, and create, rename, move or delete folders
- **Attachments** - list, save to disk, add and remove
- **Application** - check Outlook availability before doing anything else

**8 tools with 62 operations.** See [FEATURES.md](https://github.com/trsdn/mcp-server-outlook/blob/master/FEATURES.md) for the full action list.

## Requirements

- **Windows** (required - uses Outlook COM automation)
- **Classic Outlook**, installed and signed in to a profile
- **Claude Desktop** (Windows version)

> **The new Outlook will not work.** It does not expose a COM object model, so there is nothing to
> automate against. Ask Claude to check Outlook status first and it will tell you which client you
> have.

## Installation

1. Download the `.mcpb` file from the [latest release](https://github.com/trsdn/mcp-server-outlook/releases/latest)
2. Double-click to install in Claude Desktop
3. Restart Claude Desktop if prompted

That's it! Start a new conversation and ask Claude to work with Outlook.

## Usage Examples

### Example 1: Triage the inbox

**You say:** *"List the unread mail in my inbox from this week and summarize what needs a reply."*

**What happens:**
- Resolves your inbox folder
- Lists the unread messages in the requested range
- Reads the ones it needs to summarize
- Reports back which messages look like they need a response

### Example 2: Draft a reply

**You say:** *"Reply to the message I have open, saying I will get back to them on Friday with the numbers. Don't send it."*

**What happens:**
- Reads the currently open or selected message
- Creates a reply draft with the body you asked for
- Leaves the draft unsent for you to review

### Example 3: Save attachments

**You say:** *"Find the message from Accounts with the invoice and save its attachments to my Downloads folder."*

**What happens:**
- Searches mail for the message
- Lists the attachments on it
- Saves each one to the path you gave

---

**More things you can ask:**

- *"What's on my calendar tomorrow?"*
- *"Create a 30-minute appointment called Budget review on Thursday at 10."*
- *"Move everything from that sender into the Archive folder."*
- *"Mark the last five messages in my inbox as read."*
- *"Add the category Follow-up to that message."*

## Sending mail

Sending is the only irreversible thing this server does, so it is gated: a send will not happen
without explicit confirmation, and a repeated send carrying the same operation ID will not go out
twice. If a send times out, you get an *indeterminate* result rather than a false success, because
the server cannot know whether Outlook completed it.

## Tips for Best Results

- **Check status first** - ask Claude to check Outlook status if anything seems wrong; it will report whether you are on classic or new Outlook
- **Be specific** - name the folder, the sender, or the date range
- **Review before sending** - ask for a draft, read it, then ask to send it

## Privacy & Security

This server runs **entirely on your computer**. Your Outlook data:
- Never leaves your machine
- Is not sent to any external servers
- Is not used for training AI models

**Zero Logging:** This software does not collect any telemetry, usage statistics, or analytics data. No data is transmitted to external services.

## Troubleshooting

**Claude says the tool isn't available:**
- Restart Claude Desktop after installation
- Check Settings -> Integrations to verify the Outlook MCP Server is enabled

**Outlook operations fail:**
- Confirm you are running **classic** Outlook, not the new Outlook
- Ensure Outlook is running and signed in to a profile
- Some operations are blocked by the Outlook Object Model Guard; the server reports those denials distinctly rather than treating them as ordinary errors

**Need help?**
- [Report an issue](https://github.com/trsdn/mcp-server-outlook/issues)
- [Full documentation](https://github.com/trsdn/mcp-server-outlook)

## Links

- [GitHub Repository](https://github.com/trsdn/mcp-server-outlook)
- [Feature Reference](https://github.com/trsdn/mcp-server-outlook/blob/master/FEATURES.md)
- [Agent Skills](https://github.com/trsdn/mcp-server-outlook/blob/master/skills/README.md) - Cross-platform AI guidance
- [Privacy Policy](https://github.com/trsdn/mcp-server-outlook/blob/master/SECURITY.md)
- [License (MIT)](https://github.com/trsdn/mcp-server-outlook/blob/master/LICENSE)
