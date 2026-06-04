# Outlook Toolkit

SDK, CLI, and MCP server for Microsoft Outlook

A monorepo containing the SDK, CLI, and MCP server for the Outlook API.

## Packages

| Package | Description |
|---------|-------------|
| [`@outlook-toolkit/sdk`](./packages/sdk) | Core SDK with types, API client, and business logic |
| [`@outlook-toolkit/cli`](./packages/cli) | Command-line interface (Stricli) |
| [`@outlook-toolkit/mcp`](./packages/mcp) | MCP server for AI assistants (FastMCP) |

## Getting Started

```bash
# Install dependencies
bun install

# Build all packages
bun run build

# Run the CLI
bun run dev:cli -- --help

# Run the MCP server (stdio mode for Claude Desktop)
bun run dev:mcp
```

## Architecture

```
packages/sdk/     <-- Types, API client, business logic (foundation)
    ^       ^
    |       |
packages/cli/   packages/mcp/
    (Stricli)    (FastMCP)
```

Both the CLI and MCP server are thin wrappers over the SDK. If the REST API
changes, you update the SDK and both consumers get the fix automatically.

## Development

```bash
# Run tests across all packages
bun test

# Build a specific package
cd packages/sdk && bun run build
```

## Adding a New API Operation

1. Add types to `packages/sdk/src/types.ts`
2. Add the client method to `packages/sdk/src/client.ts`
3. Add a CLI command in `packages/cli/src/commands/`
4. Add an MCP tool in `packages/mcp/src/tools/`
