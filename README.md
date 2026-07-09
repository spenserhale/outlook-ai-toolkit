# Outlook Toolkit

SDK, CLI, and MCP server for Microsoft Outlook via Microsoft Graph.

Supports personal accounts (outlook.com / MSA) and work/school accounts (Entra ID / Microsoft 365). Uses delegated OAuth 2.0 with PKCE — no client secret, tokens stored encrypted in your OS keychain.

## Packages

| Package | Description |
|---------|-------------|
| [`@outlook-toolkit/sdk`](./packages/sdk) | Core SDK with types, auth, Graph client, and mail operations |
| [`@outlook-toolkit/cli`](./packages/cli) | Command-line interface (Stricli) |
| [`@outlook-toolkit/mcp`](./packages/mcp) | MCP server for AI assistants (FastMCP) |

## Install the CLI

### Recommended: standalone binary

No Node.js, no npm, no PATH conflicts. One file.

**macOS and Linux:**

```sh
curl -fsSL https://raw.githubusercontent.com/spenserhale/outlook-ai-toolkit/main/scripts/install.sh | sh
```

The script detects your OS + architecture, downloads the matching binary from
the [latest release](https://github.com/spenserhale/outlook-ai-toolkit/releases/latest),
verifies its SHA256, and installs to `$HOME/.local/bin/outlook`. Pin a specific
version with `OUTLOOK_TOOLKIT_VERSION=v0.1.1` or change the install directory
with `OUTLOOK_TOOLKIT_INSTALL=$HOME/bin`.

**Windows:** download `outlook-windows-x64.exe` from the
[latest release](https://github.com/spenserhale/outlook-ai-toolkit/releases/latest)
and put it on your `PATH`.

**Updating:** re-run the install command, or use the built-in:

```sh
outlook upgrade          # install latest
outlook upgrade --check  # check without installing
```

Available binaries: `outlook-linux-{x64,arm64}`, `outlook-darwin-{x64,arm64}`,
`outlook-windows-x64.exe`. A `.sha256` sits next to each one; an aggregated
`SHASUMS256.txt` is attached to the release.

After install, run `outlook --help` or `outlook agent-context --json` to see
every command. Use `outlook` directly instead of `bun run dev:cli --` in the
examples below (e.g. `outlook mail list`, `outlook auth login`).

### Alternative: build from source

```sh
git clone https://github.com/spenserhale/outlook-ai-toolkit.git
cd outlook-ai-toolkit
bun install
bun run build
```

Before you can sign in, you still need a one-time [Azure App
Registration](#azure-app-registration) to get a `client_id`.

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
4. **Redirect URI:** Platform = **Mobile and desktop applications**, URI = `http://localhost/callback`
   - The `/callback` path is required — it's where the CLI's local callback server listens. Azure ignores the port for loopback URIs on this platform, so the toolkit's random port is fine, but the path must match exactly.
5. After creating, copy the **Application (client) ID**
6. Go to **Authentication** → enable **Allow public client flows** → Save
7. Go to **API permissions** → Add the following **Microsoft Graph delegated** permissions:
   - `Mail.Read`
   - `Mail.ReadWrite`
   - `Mail.Send`
   - `User.Read`
   - `offline_access` (usually pre-added)

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

# Machine-readable CLI schema (for agents)
bun run dev:cli -- agent-context
```

### Output formats

`mail` commands render [TOON](https://github.com/toon-format/toon) by default (compact, token-efficient — ideal for agents). Pick another format with a flag:

```bash
bun run dev:cli -- mail list            # TOON (default)
bun run dev:cli -- mail list --json     # JSON
bun run dev:cli -- mail list --csv      # CSV (mail list only)
```

### Message body control

`mail list` and `mail get` can shape how message bodies are returned:

```bash
# How much body to include (mail list only): none | preview | full (default: preview)
bun run dev:cli -- mail list --body=full

# Body format: text | markdown | html (default: text)
bun run dev:cli -- mail get <id> --bodyFormat=markdown
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
bun run dev:cli -- profile save personal --clientId=abc123 --tenantId=consumers
bun run dev:cli -- profile save work --clientId=xyz789 --tenantId=<your-tenant-guid>

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
