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
examples below (e.g. `outlook mail list`, `outlook login`).

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

- A registered Azure application (see below) — needed to get a `client_id`
- [Bun](https://bun.sh) ≥ 1.0 — only if building from source (the standalone binary needs nothing)

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

## Getting started

**Run this first:**

```sh
outlook login
```

On the first run it walks you through a one-time setup: it asks for your Azure
app **Client ID** (see [Azure App Registration](#azure-app-registration) to
create one — ~2 min), saves it as your default profile, and opens the browser
for Microsoft sign-in. Tokens are stored in your OS keychain (macOS Keychain /
Windows Credential Store / Linux Secret Service). After that, `outlook login`
just signs you in.

```sh
outlook status     # who am I signed in as?
outlook logout     # sign out and clear tokens
```

> Running from source instead of the installed binary? Use `bun run dev:cli --`
> in place of `outlook` (e.g. `bun run dev:cli -- login`).

### Common commands

```sh
# List inbox (default: 25 messages)
outlook mail list

# Get a specific message
outlook mail get <id>

# Send email
outlook mail send --to=someone@example.com --subject="Hello" --body="<p>Hi</p>"

# Reply to a message
outlook mail reply <id> --body="<p>Thanks</p>"

# Create a draft
outlook mail draft --to=someone@example.com --subject="Draft" --body="<p>Draft body</p>"

# Paginate
outlook mail list --limit=50 --cursor=<nextLink>

# Machine-readable CLI schema (for agents)
outlook agent-context --json
```

### Output formats

`mail` commands render [TOON](https://github.com/toon-format/toon) by default (compact, token-efficient — ideal for agents). Pick another format with a flag:

```sh
outlook mail list            # TOON (default)
outlook mail list --json     # JSON
outlook mail list --csv      # CSV (mail list only)
```

### Message body control

`mail list` and `mail get` can shape how message bodies are returned:

```sh
# How much body to include (mail list only): none | preview | full (default: preview)
outlook mail list --body=full

# Body format: text | markdown | html (default: text)
outlook mail get <id> --bodyFormat=markdown
```

### Configuration via environment (optional)

Instead of the guided `outlook login` setup, you can supply credentials through
the environment or a `.env` file — handy for scripting and CI:

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

With those set, `outlook login` skips the prompt and signs you straight in.

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

**Before using MCP tools**, sign in once via the CLI:

```sh
outlook login
```

The MCP server reads the saved tokens — it cannot open a browser itself. If a tool reports "not authenticated", run `outlook login` from the CLI and try again.

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

Your first `outlook login` creates a `default` profile. To add more accounts,
use `profile add` — it saves the profile and signs in, prompting for the Client
ID if you don't pass `--clientId`:

```sh
# Add named accounts (each opens a browser sign-in)
outlook profile add personal --clientId=abc123 --tenantId=consumers
outlook profile add work --clientId=xyz789 --tenantId=<your-tenant-guid>

outlook profile list                 # see saved profiles

# Target a profile with --profile
outlook mail list --profile=personal
outlook mail list --profile=work
outlook login --profile=work         # re-authenticate a profile
```

Set `OUTLOOK_PROFILE=work` to make a profile the default for a shell session.
(`profile save` is the non-interactive variant — it stores a profile without
signing in, for scripting.)

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
| 5 | Not authenticated — run `outlook login` |
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
