# Power BI MCP for Claude

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.x-3178c6.svg)](https://www.typescriptlang.org/)
[![MCP](https://img.shields.io/badge/MCP-Power%20BI%20%2B%20Fabric-111827.svg)](https://modelcontextprotocol.io/)

Claude Desktop and Claude Code MCP server for Microsoft Power BI and Fabric. It supports Microsoft device-code sign-in, Power BI workspace and semantic model discovery, active semantic model context, DAX query generation/execution, schema metadata, report metadata, and redacted diagnostics.

Local Node/TypeScript MCP proxy for Claude. It forwards JSON-RPC messages to Microsoft's remote Power BI MCP endpoint, exposes stable local wrapper tools for Claude discovery, and acquires delegated Microsoft Entra ID tokens with MSAL device-code authentication.

The key difference from a terminal-only proxy is that first-time sign-in can happen from Claude chat. If Claude advertises MCP URL-mode elicitation, the server uses it. Claude Desktop currently starts this extension with MCP Apps UI capabilities instead, so the server also exposes local auth tools that return the Microsoft URL and device code directly in chat.

Remote endpoint:

```text
https://api.fabric.microsoft.com/v1/mcp/powerbi
```

Power BI data access goes through this Fabric MCP endpoint. The local tools only handle sign-in and provide stable wrapper names for Claude Desktop to call.

## Features

- Claude Desktop MCPB extension and Claude Code plugin support.
- Microsoft Entra ID device-code authentication through MSAL.
- Power BI workspace and semantic model discovery wrappers.
- Default and chat-scoped workspace/semantic model context.
- Fabric MCP wrappers for semantic model schema, DAX query generation, DAX execution, and report metadata.
- Redacted diagnostics for auth, context, local wrapper tools, and upstream Fabric MCP tool discovery.
- Fabric-only routing through `https://api.fabric.microsoft.com/v1/mcp/powerbi`.

## Install

```bash
npm install
npm run build
```

## Commands

```bash
npm run start          # run stdio MCP proxy
npm run start:http     # run hosted Streamable HTTP-style MCP server
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

The root `manifest.json` describes the Claude Desktop extension. It passes `POWERBI_MCP_CACHE_DIR`, `POWERBI_MCP_TENANT_ID`, `POWERBI_MCP_CLIENT_ID`, and optional default workspace/semantic-model settings into the server.

## Claude Code Plugin

The `.claude-plugin/plugin.json`, `.mcp.json`, and `skills/powerbi/SKILL.md` files let the same server run as a Claude Code plugin:

```bash
npm run build
claude --plugin-dir /path/to/Power-BI-MCP-Node
```

## Hosted Remote MCP

Run the hosted HTTP entrypoint when you need a remote MCP endpoint for Claude web/mobile custom connectors or a gateway such as LiteLLM:

```bash
npm run build
PORT=3000 POWERBI_MCP_CACHE_DIR=/data/powerbi-mcp-claude npm run start:http
```

The hosted server exposes:

- `POST /mcp` for MCP JSON-RPC requests.
- `GET /healthz` for container health checks.

Hosted mode creates a separate MSAL token cache per `Mcp-Session-Id` under:

```text
${POWERBI_MCP_CACHE_DIR}/sessions/<session-id>/msal_token_cache.json
```

This prevents one Claude MCP session from reusing another session's Power BI token. The server returns a strong random `Mcp-Session-Id` header and expires inactive sessions after `POWERBI_MCP_HTTP_SESSION_TTL_SECONDS` seconds, defaulting to two hours.

Docker:

```bash
docker build -t powerbi-mcp-claude .
docker run --rm -p 3000:3000 \
  -v powerbi-mcp-cache:/data/powerbi-mcp-claude \
  -e POWERBI_MCP_TENANT_ID=organizations \
  powerbi-mcp-claude
```

For production, put the container behind LiteLLM, an ingress, firewall rules, or another network boundary. Anthropic IP allowlisting limits where connector traffic can originate, but it does not identify which Claude user or organization initiated the request. Per-session Power BI auth avoids a shared server-wide Power BI token cache; full OAuth connector auth can be added later if stronger user identity is required at the MCP endpoint.

## Authentication UX

The server first tries silent token acquisition from the local MSAL cache. It always exposes these local tools so Claude can discover them at connector startup:

- `powerbi_auth_start`
- `powerbi_auth_status`
- `powerbi_auth_logout`
- `powerbi_diagnostics`
- `powerbi_get_default_context`
- `powerbi_set_context`
- `powerbi_clear_context`
- `powerbi_list_workspaces`
- `powerbi_list_semantic_models`
- `powerbi_get_semantic_model_schema`
- `powerbi_generate_dax_query`
- `powerbi_execute_dax_query`
- `powerbi_get_report_metadata`

If no cached token is available:

1. Claude calls `powerbi_auth_start`.
2. `powerbi_auth_start` returns Microsoft's verification URL and a device code in chat.
3. The user completes Microsoft sign-in in the browser.
4. MSAL stores refreshable tokens in the configured cache directory.
5. Later requests, including workspace and semantic-model listing, refresh silently when possible and are sent to the Fabric MCP endpoint.

If chat-based sign-in is not convenient, run `npm run login` once from a normal terminal, then reconnect the MCP server.

## Diagnostics

Use `powerbi_diagnostics` when Claude needs to verify whether the extension is healthy. It returns redacted status for the server version, authentication state, active/default context presence, local wrapper tools, and upstream Fabric MCP query tool discovery.

Diagnostics output intentionally omits tokens, auth headers, account email, cache paths, and raw workspace or semantic model IDs.

## Default Power BI Context

Some Fabric MCP tools require a workspace ID or semantic model ID even when the endpoint does not advertise a tool that can list every available ID. In Claude Desktop settings, optionally configure one default workspace and one default semantic model:

- Default workspace ID
- Default workspace name
- Default semantic model ID
- Default semantic model name

The extension uses these defaults only when a Fabric MCP tool requires a missing ID. Explicit IDs supplied in chat are preserved.

You can also switch the active context inside a chat without changing Claude Desktop settings:

- `powerbi_set_context` sets a chat-session workspace ID, semantic model ID, or display names.
- `powerbi_clear_context` clears the chat-session override and returns to Claude Desktop settings.
- `powerbi_get_default_context` shows the active context, chat override, and configured defaults.

Context precedence is: explicit IDs in the current tool call, then chat-session context, then Claude Desktop settings.

The query tools are local wrappers around Fabric MCP tools so Claude can discover them before sign-in:

- `powerbi_get_semantic_model_schema` forwards to Fabric MCP `GetSemanticModelSchema`.
- `powerbi_generate_dax_query` forwards to Fabric MCP `GenerateQuery`.
- `powerbi_execute_dax_query` forwards to Fabric MCP `ExecuteQuery`.
- `powerbi_get_report_metadata` forwards to Fabric MCP `GetReportMetadata`.

The wrappers still require Microsoft sign-in and still route all Power BI operations through `https://api.fabric.microsoft.com/v1/mcp/powerbi`.

## Environment Variables

| Variable | Default |
| --- | --- |
| `POWERBI_MCP_URL` | `https://api.fabric.microsoft.com/v1/mcp/powerbi` |
| `POWERBI_MCP_CLIENT_ID` | built-in public Power BI client ID |
| `POWERBI_MCP_TENANT_ID` | `organizations` |
| `POWERBI_MCP_SCOPES` | `https://analysis.windows.net/powerbi/api/.default` |
| `POWERBI_MCP_DEFAULT_WORKSPACE_ID` | empty |
| `POWERBI_MCP_DEFAULT_WORKSPACE_NAME` | empty |
| `POWERBI_MCP_DEFAULT_SEMANTIC_MODEL_ID` | empty |
| `POWERBI_MCP_DEFAULT_SEMANTIC_MODEL_NAME` | empty |
| `POWERBI_MCP_CACHE_DIR` | `~/.powerbi-mcp-claude`, or `${CLAUDE_PLUGIN_DATA}/cache` in Claude Code plugin mode |
| `POWERBI_MCP_TIMEOUT_SECONDS` | `60` |
| `POWERBI_MCP_DEVICE_CODE_TIMEOUT_SECONDS` | `900` |
| `POWERBI_MCP_ALLOW_INTERACTIVE_AUTH` | `true` |
| `PORT` / `POWERBI_MCP_HTTP_PORT` | `3000` in hosted HTTP mode |
| `POWERBI_MCP_HTTP_PATH` | `/mcp` |
| `POWERBI_MCP_HTTP_SESSION_TTL_SECONDS` | `7200` |

Do not paste Microsoft access tokens or refresh tokens into Claude. Use device-code sign-in or a tenant-approved public-client app registration.
