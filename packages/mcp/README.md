# @outlook-toolkit/mcp

MCP server for Outlook, built with [FastMCP](https://github.com/punkpeye/fastmcp).

## Tools

| Tool | Description |
|------|-------------|
| `list_resources` | List resources with pagination |
| `get_resource` | Get a resource by ID |
| `create_resource` | Create a new resource |
| `delete_resource` | Delete a resource |

## Setup with Claude Desktop

Add this to your Claude Desktop config (`~/Library/Application Support/Claude/claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "outlook-toolkit": {
      "command": "bun",
      "args": ["run", "/Users/spenser/Code/Toolkits/outlook-toolkit/packages/mcp/src/index.ts"],
      "env": {
        "OUTLOOK_API_KEY": "your-api-key-here"
      }
    }
  }
}
```

## Development

```bash
# Run in stdio mode
bun run dev

# Inspect with FastMCP inspector
bun run inspect
```
