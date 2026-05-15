# Power BI MCP for Claude

Local Node/TypeScript MCP proxy for Claude. It forwards JSON-RPC messages to Microsoft's remote Power BI MCP endpoint, exposes stable local wrapper tools for Claude discovery, and acquires delegated Microsoft Entra ID tokens with MSAL device-code authentication.

The key difference from a terminal-only proxy is that first-time sign-in can happen from Claude chat. If Claude advertises MCP URL-mode elicitation, the server uses it. Claude Desktop currently starts this extension with MCP Apps UI capabilities instead, so the server also exposes local auth tools that return the Microsoft URL and device code directly in chat.

Remote endpoint:

```text
https://api.fabric.microsoft.com/v1/mcp/powerbi
```

Power BI data access goes through this Fabric MCP endpoint. The local tools only handle sign-in and provide stable wrapper names for Claude Desktop to call.

## Install

```bash
npm install
npm run build
```

## Commands

```bash
npm run start          # run stdio MCP proxy
npm run login          # authenticate in a terminal and populate the cache
npm run status         # show non-sensitive cache/account info
npm run logout         # delete the local MSAL token cache
npm test               # run focused transport tests
```

## Claude Desktop

Package as an MCPB desktop extension after building:

```bash
npm run build
npm run pack:mcpb
```

The root `manifest.json` describes the Claude Desktop extension. It passes `POWERBI_MCP_CACHE_DIR`, `POWERBI_MCP_TENANT_ID`, and `POWERBI_MCP_CLIENT_ID` into the server.

## Claude Code Plugin

The `.claude-plugin/plugin.json`, `.mcp.json`, and `skills/powerbi/SKILL.md` files let the same server run as a Claude Code plugin:

```bash
npm run build
claude --plugin-dir /Users/robertomoreno/Documents/Power-BI-MCP-Node
```

## Authentication UX

The server first tries silent token acquisition from the local MSAL cache. It always exposes these local tools so Claude can discover them at connector startup:

- `powerbi_auth_start`
- `powerbi_auth_status`
- `powerbi_auth_logout`
- `powerbi_list_workspaces`
- `powerbi_list_semantic_models`

If no cached token is available:

1. Claude calls `powerbi_auth_start`.
2. `powerbi_auth_start` returns Microsoft's verification URL and a device code in chat.
3. The user completes Microsoft sign-in in the browser.
4. MSAL stores refreshable tokens in the configured cache directory.
5. Later requests, including workspace and semantic-model listing, refresh silently when possible and are sent to the Fabric MCP endpoint.

If chat-based sign-in is not convenient, run `npm run login` once from a normal terminal, then reconnect the MCP server.

## Environment Variables

| Variable | Default |
| --- | --- |
| `POWERBI_MCP_URL` | `https://api.fabric.microsoft.com/v1/mcp/powerbi` |
| `POWERBI_MCP_CLIENT_ID` | built-in public Power BI client ID |
| `POWERBI_MCP_TENANT_ID` | `organizations` |
| `POWERBI_MCP_SCOPES` | `https://analysis.windows.net/powerbi/api/.default` |
| `POWERBI_MCP_CACHE_DIR` | `~/.powerbi-mcp-claude`, or `${CLAUDE_PLUGIN_DATA}/cache` in Claude Code plugin mode |
| `POWERBI_MCP_TIMEOUT_SECONDS` | `60` |
| `POWERBI_MCP_DEVICE_CODE_TIMEOUT_SECONDS` | `900` |
| `POWERBI_MCP_ALLOW_INTERACTIVE_AUTH` | `true` |

Do not paste Microsoft access tokens or refresh tokens into Claude. Use device-code sign-in or a tenant-approved public-client app registration.
