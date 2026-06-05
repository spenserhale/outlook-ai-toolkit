# Outlook Toolkit

SDK, CLI, and MCP server for Microsoft Outlook via Microsoft Graph.

Supports personal accounts (outlook.com / MSA) and work/school accounts (Entra ID / Microsoft 365). Uses delegated OAuth 2.0 with PKCE — no client secret, tokens stored encrypted in your OS keychain.

---

## Prerequisites

- [Bun](https://bun.sh) ≥ 1.0
- A registered Azure application (see below)

---

## Azure App Registration

You need a Microsoft Entra app registration to get a `client_id`. This is a one-time setup per account type.

1. Go to [portal.azure.com](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** → **New registration**
2. Name it anything (e.g. `outlook-toolkit`)
3. **Supported account types:**
   - Personal accounts only → *Personal Microsoft accounts only*
   - Work/school + personal → *Accounts in any organizational directory and personal Microsoft accounts*
   - Work/school only → *Accounts in this organizational directory only*
4. **Redirect URI:** Platform = **Mobile and desktop applications**, URI = `http://localhost`
5. After creating, copy the **Application (client) ID**
6. Go to **Authentication** → enable **Allow public client flows** → Save
7. Go to **API permissions** → Add the following **Microsoft Graph delegated** permissions:
   - `Mail.Read`
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `User.Read`
   - `offline_access` (usually pre-added)

---

## Installation

```bash
git clone <repo>
cd outlook-toolkit
bun install
bun run build
```

---

## Configuration

Copy `.env.example` to `.env` and fill in your values:

```bash
cp .env.example .env
```

```
OUTLOOK_CLIENT_ID=your-azure-app-client-id
OUTLOOK_TENANT_ID=consumers
```

`OUTLOOK_TENANT_ID` controls which account types are allowed:

| Value | Use for |
|-------|---------|
| `consumers` | Personal outlook.com / MSA accounts |
| `common` | Personal + work/school (multi-tenant) |
| `<guid>` | A specific Entra ID tenant (work/school only) |

---

## CLI Setup

### Authenticate

```bash
bun run dev:cli -- auth login
```

This opens your browser for the Microsoft sign-in flow. After completing sign-in, tokens are saved to your OS keychain (macOS Keychain / Windows Credential Store / Linux Secret Service).

### Verify

```bash
bun run dev:cli -- auth status
```

### Common commands

```bash
# List inbox (default: 25 messages)
bun run dev:cli -- mail list

# Get a specific message
bun run dev:cli -- mail get <id>

# Send email
bun run dev:cli -- mail send --to=someone@example.com --subject="Hello" --body="<p>Hi</p>"

# Reply to a message
bun run dev:cli -- mail reply <id> --body="<p>Thanks</p>"

# Create a draft
bun run dev:cli -- mail draft --to=someone@example.com --subject="Draft" --body="<p>Draft body</p>"

# Paginate
bun run dev:cli -- mail list --limit=50 --cursor=<nextLink>

# JSON output
bun run dev:cli -- mail list --json

# Machine-readable CLI schema (for agents)
bun run dev:cli -- agent-context
```

### Sign out

```bash
bun run dev:cli -- auth logout
```

---

## MCP Setup (Claude Desktop)

Add the MCP server to your Claude Desktop config at `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "outlook": {
      "command": "bun",
      "args": ["run", "/absolute/path/to/outlook-toolkit/packages/mcp/src/index.ts"],
      "env": {
        "OUTLOOK_CLIENT_ID": "your-azure-app-client-id",
        "OUTLOOK_TENANT_ID": "consumers"
      }
    }
  }
}
```

Replace `/absolute/path/to/outlook-toolkit` with the actual path on your machine.

**Before using MCP tools**, authenticate once via the CLI:

```bash
bun run dev:cli -- auth login
```

The MCP server reads the saved tokens — it cannot open a browser itself. If a tool reports "not authenticated", run `auth login` from the CLI and try again.

### Available MCP tools

| Tool | What it does |
|------|-------------|
| `outlook_auth_status` | Check authentication state |
| `outlook_auth_logout` | Clear saved tokens |
| `outlook_mail_list` | List messages with optional filter/limit/cursor |
| `outlook_mail_get` | Get a single message by ID |
| `outlook_mail_send` | Send an email |
| `outlook_mail_reply` | Reply to a message thread |
| `outlook_mail_create_draft` | Create a draft without sending |
| `outlook_mail_sync` | Delta sync — returns only changed messages |

---

## Multiple Accounts (Profiles)

To switch between personal and work accounts, save named profiles:

```bash
# Save profiles
bun run dev:cli -- profile save personal --client-id=abc123 --tenant-id=consumers
bun run dev:cli -- profile save work --client-id=xyz789 --tenant-id=<your-tenant-guid>

# Use a profile
bun run dev:cli -- --profile=personal auth login
bun run dev:cli -- --profile=work auth login

bun run dev:cli -- --profile=personal mail list
bun run dev:cli -- --profile=work mail list
```

For MCP, run separate server instances with different env vars:

```json
{
  "mcpServers": {
    "outlook-personal": {
      "command": "bun",
      "args": ["run", "/absolute/path/to/outlook-toolkit/packages/mcp/src/index.ts"],
      "env": {
        "OUTLOOK_CLIENT_ID": "abc123",
        "OUTLOOK_TENANT_ID": "consumers"
      }
    },
    "outlook-work": {
      "command": "bun",
      "args": ["run", "/absolute/path/to/outlook-toolkit/packages/mcp/src/index.ts"],
      "env": {
        "OUTLOOK_CLIENT_ID": "xyz789",
        "OUTLOOK_TENANT_ID": "<your-tenant-guid>"
      }
    }
  }
}
```

Each instance maintains its own token cache keyed by `client_id`.

---

## Exit Codes

| Code | Meaning |
|------|---------|
| 0 | Success |
| 1 | Network / fetch error |
| 2 | Invalid flag or argument |
| 3 | Missing config (`OUTLOOK_CLIENT_ID` / `OUTLOOK_TENANT_ID`) |
| 4 | Not found |
| 5 | Not authenticated — run `auth login` |
| 6 | Rate limited |

---

## Architecture

```
packages/sdk/     Types, auth, Graph client, mail operations (foundation)
    ^       ^
    |       |
packages/cli/   packages/mcp/
  (Stricli)      (FastMCP stdio)
```

Both CLI and MCP are thin wrappers over the SDK. All API logic lives in the SDK.

---

## Development

```bash
bun test            # run all tests
bun run lint        # type-check all packages
bun run build       # build all packages
bun run clean       # remove dist + node_modules
```
