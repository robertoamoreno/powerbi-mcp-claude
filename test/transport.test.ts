import assert from "node:assert/strict";
import { test } from "node:test";
import type { DeviceCodeResponse } from "@azure/msal-common/node";

import { normalizeId, type JSONObject, type JSONRPCPayload, type JSONValue } from "../src/jsonrpc.js";
import { LocalPowerBITools } from "../src/localTools.js";
import type { RemoteForwarder } from "../src/remote.js";
import { parseSsePayload, RemoteMCPError } from "../src/remote.js";
import { applyDefaultArguments, type RemoteTool } from "../src/remotePowerBITools.js";
import { MCPProxyTransport } from "../src/transport.js";

const deviceCode: DeviceCodeResponse = {
  userCode: "ABCD-EFGH",
  deviceCode: "secret-device-code",
  verificationUri: "https://www.microsoft.com/link",
  expiresIn: 900,
  interval: 5,
  message: "To sign in, use a web browser to open https://www.microsoft.com/link and enter ABCD-EFGH.",
};

test("device-code prompt uses URL-mode elicitation when the client supports it", async () => {
  const written: JSONValue[] = [];
  const remote: RemoteForwarder = {
    async forward(): Promise<JSONValue> {
      return null;
    },
  };
  const transport = new MCPProxyTransport(remote, (payload) => written.push(payload));

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 1,
    method: "initialize",
    params: {
      protocolVersion: "2025-11-25",
      capabilities: {
        elicitation: {
          form: {},
          url: {},
        },
      },
    },
  });

  assert.equal(transport.supportsDeviceCodePrompt(), true);

  const prompt = transport.promptDeviceCode(deviceCode);
  const elicitation = findElicitation(written);
  const params = asObject(elicitation.params);

  assert.ok(elicitation);
  assert.equal(params.mode, "url");
  assert.equal(params.url, "https://www.microsoft.com/link");
  assert.match(String(params.message), /ABCD-EFGH/);

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: elicitation.id,
    result: {
      action: "accept",
    },
  });

  assert.equal(await prompt, "accepted");
});

test("elicitation responses are consumed locally and not forwarded to Power BI", async () => {
  const written: JSONValue[] = [];
  const forwarded: JSONRPCPayload[] = [];
  const remote: RemoteForwarder = {
    async forward(payload: JSONRPCPayload): Promise<JSONValue> {
      forwarded.push(payload);
      return null;
    },
  };
  const transport = new MCPProxyTransport(remote, (payload) => written.push(payload));

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 1,
    method: "initialize",
    params: {
      capabilities: {
        elicitation: {
          url: {},
        },
      },
    },
  });

  const prompt = transport.promptDeviceCode(deviceCode);
  const elicitation = findElicitation(written);
  assert.ok(elicitation);

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: elicitation.id,
    result: {
      action: "decline",
    },
  });

  assert.equal(await prompt, "declined");
  assert.equal(forwarded.length, 0);
});

test("initialize is answered locally so missing auth does not disconnect Claude Desktop", () => {
  const written: JSONValue[] = [];
  const forwarded: JSONRPCPayload[] = [];
  const transport = new MCPProxyTransport(
    {
      async forward(payload: JSONRPCPayload): Promise<JSONValue> {
        forwarded.push(payload);
        return null;
      },
    },
    (payload) => written.push(payload),
  );

  transport.handleIncoming({
    method: "initialize",
    params: {
      protocolVersion: "2025-11-25",
      capabilities: {
        extensions: {
          "io.modelcontextprotocol/ui": {
            mimeTypes: ["text/html;profile=mcp-app"],
          },
        },
      },
      clientInfo: {
        name: "claude-ai",
        version: "0.1.0",
      },
    },
    jsonrpc: "2.0",
    id: 0,
  });

  assert.equal(forwarded.length, 0);
  assert.equal(written.length, 1);
  assert.equal(isJsonObject(written[0]) && written[0].id, 0);
  assert.equal(isJsonObject(written[0]) && isJsonObject(written[0].result) && isJsonObject(written[0].result.capabilities), true);
});

test("tools/list exposes Power BI domain tools before authentication", async () => {
  const written: JSONValue[] = [];
  const forwarded: JSONRPCPayload[] = [];
  const localTools = new LocalPowerBITools({
    async getCachedAccessToken(): Promise<undefined> {
      return undefined;
    },
  } as never);

  const transport = new MCPProxyTransport(
    {
      async forward(payload: JSONRPCPayload): Promise<JSONValue> {
        forwarded.push(payload);
        return null;
      },
    },
    (payload) => written.push(payload),
    localTools,
  );

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 2,
    method: "tools/list",
  });

  await new Promise((resolve) => setImmediate(resolve));

  assert.equal(forwarded.length, 0);
  const response = asObject(written[0]);
  const result = asObject(response.result);
  assert.ok(Array.isArray(result.tools));
  const toolNames = result.tools.map((tool) => asObject(tool).name);
  assert.ok(toolNames.includes("powerbi_auth_start"));
  assert.ok(toolNames.includes("powerbi_diagnostics"));
  assert.ok(toolNames.includes("powerbi_get_default_context"));
  assert.ok(toolNames.includes("powerbi_set_context"));
  assert.ok(toolNames.includes("powerbi_clear_context"));
  assert.ok(toolNames.includes("powerbi_list_workspaces"));
  assert.ok(toolNames.includes("powerbi_list_semantic_models"));
  assert.ok(toolNames.includes("powerbi_get_semantic_model_schema"));
  assert.ok(toolNames.includes("powerbi_generate_dax_query"));
  assert.ok(toolNames.includes("powerbi_execute_dax_query"));
  assert.ok(toolNames.includes("powerbi_get_report_metadata"));
});

test("diagnostics are available before authentication and do not call upstream tools", async () => {
  const written: JSONValue[] = [];
  const forwarded: JSONRPCPayload[] = [];
  const localTools = new LocalPowerBITools({
    async getCachedAccessToken(): Promise<undefined> {
      return undefined;
    },
    async deviceLoginStatus() {
      return {
        status: "not_started",
        authenticated: false,
        message: "Microsoft authentication has not been started.",
      };
    },
    async tokenInfo() {
      return {
        cachePath: "/tmp/powerbi-mcp-claude/msal_token_cache.json",
        cacheExists: false,
        accountCount: 0,
        usernames: [],
      };
    },
  } as never);

  const transport = new MCPProxyTransport(
    {
      async forward(payload: JSONRPCPayload): Promise<JSONValue> {
        forwarded.push(payload);
        throw new Error("Diagnostics should not call upstream tools before auth.");
      },
    },
    (payload) => written.push(payload),
    localTools,
  );

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 0,
    method: "initialize",
    params: {
      protocolVersion: "2025-11-25",
      capabilities: {},
    },
  });
  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 20,
    method: "tools/call",
    params: {
      name: "powerbi_diagnostics",
      arguments: {},
    },
  });

  await new Promise((resolve) => setImmediate(resolve));

  assert.equal(forwarded.length, 0);
  const response = asObject(written.find((payload) => isJsonObject(payload) && payload.id === 20));
  const result = asObject(response.result);
  const diagnostics = diagnosticsFromResult(result);
  const auth = asObject(diagnostics.auth);
  const upstream = asObject(diagnostics.upstream);
  assert.equal(auth.status, "not_started");
  assert.equal(auth.authenticated, false);
  assert.equal(upstream.status, "skipped_auth_required");
  assert.doesNotMatch(JSON.stringify(response), /user@example\.com|msal_token_cache|Bearer/i);
});

test("diagnostics report upstream Fabric query tool availability after authentication", async () => {
  const written: JSONValue[] = [];
  const forwarded: JSONRPCPayload[] = [];
  const workspaceId = "00000000-0000-0000-0000-000000000001";
  const semanticModelId = "00000000-0000-0000-0000-000000000002";
  const localTools = new LocalPowerBITools(
    {
      async getCachedAccessToken(): Promise<string> {
        return "cached-token";
      },
      async deviceLoginStatus() {
        return {
          status: "authenticated",
          authenticated: true,
          username: "user.com",
          message: "Authenticated with Microsoft as user.com.",
        };
      },
      async tokenInfo() {
        return {
          cachePath: "/tmp/powerbi-mcp-claude/msal_token_cache.json",
          cacheExists: true,
          accountCount: 1,
          usernames: ["user.com"],
        };
      },
    } as never,
    {
      remoteUrl: "https://api.fabric.microsoft.com/v1/mcp/powerbi",
      defaultWorkspaceId: workspaceId,
      defaultSemanticModelId: semanticModelId,
    },
  );

  const transport = new MCPProxyTransport(
    {
      async forward(payload: JSONRPCPayload): Promise<JSONValue> {
        forwarded.push(payload);
        if (!Array.isArray(payload) && payload.method === "initialize") {
          return {
            jsonrpc: "2.0",
            id: normalizeId(payload.id),
            result: {
              capabilities: {
                tools: {},
              },
            },
          };
        }
        if (!Array.isArray(payload) && payload.method === "tools/list") {
          return queryTools(payload);
        }
        return null;
      },
    },
    (payload) => written.push(payload),
    localTools,
  );

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 0,
    method: "initialize",
    params: {
      protocolVersion: "2025-11-25",
      capabilities: {},
    },
  });
  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 21,
    method: "tools/call",
    params: {
      name: "powerbi_diagnostics",
      arguments: {},
    },
  });

  await new Promise((resolve) => setImmediate(resolve));

  assert.equal(
    forwarded.some((payload) => !Array.isArray(payload) && payload.method === "tools/list"),
    true,
  );
  const response = asObject(written.find((payload) => isJsonObject(payload) && payload.id === 21));
  const result = asObject(response.result);
  const diagnostics = diagnosticsFromResult(result);
  const auth = asObject(diagnostics.auth);
  const context = asObject(diagnostics.context);
  const endpoint = asObject(diagnostics.endpoint);
  const upstream = asObject(diagnostics.upstream);
  const capabilities = asObject(upstream.expectedCapabilities);
  assert.equal(auth.authenticated, true);
  assert.equal(endpoint.looksLikeFabricPowerBI, true);
  assert.equal(context.activeWorkspaceConfigured, true);
  assert.equal(context.activeSemanticModelConfigured, true);
  assert.equal(upstream.status, "ok");
  assert.equal(capabilities.ExecuteQuery, true);
  assert.equal(capabilities.GetSemanticModelSchema, true);
  assert.equal(capabilities.GenerateQuery, true);
  assert.equal(capabilities.GetReportMetadata, true);
  const serialized = JSON.stringify(response);
  assert.doesNotMatch(serialized, new RegExp(workspaceId));
  assert.doesNotMatch(serialized, new RegExp(semanticModelId));
  assert.doesNotMatch(serialized, /user@example\.com|msal_token_cache|cached-token|Bearer/i);
});

test("diagnostics report upstream discovery failures without leaking details", async () => {
  const written: JSONValue[] = [];
  const localTools = new LocalPowerBITools(
    {
      async getCachedAccessToken(): Promise<string> {
        return "cached-token";
      },
      async deviceLoginStatus() {
        return {
          status: "authenticated",
          authenticated: true,
          username: "user.com",
          message: "Authenticated with Microsoft as user.com.",
        };
      },
      async tokenInfo() {
        return {
          cachePath: "/tmp/powerbi-mcp-claude/msal_token_cache.json",
          cacheExists: true,
          accountCount: 1,
          usernames: ["user.com"],
        };
      },
    } as never,
    {
      remoteUrl: "https://api.fabric.microsoft.com/v1/mcp/powerbi",
      defaultSemanticModelId: "00000000-0000-0000-0000-000000000002",
    },
  );

  const transport = new MCPProxyTransport(
    {
      async forward(payload: JSONRPCPayload): Promise<JSONValue> {
        if (!Array.isArray(payload) && payload.method === "initialize") {
          return {
            jsonrpc: "2.0",
            id: normalizeId(payload.id),
            result: {
              capabilities: {
                tools: {},
              },
            },
          };
        }
        if (!Array.isArray(payload) && payload.method === "tools/list") {
          throw new RemoteMCPError(
            "sensitive user.com /tmp/powerbi-mcp-claude/msal_token_cache.json",
            -32000,
            503,
          );
        }
        return null;
      },
    },
    (payload) => written.push(payload),
    localTools,
  );

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 0,
    method: "initialize",
    params: {
      protocolVersion: "2025-11-25",
      capabilities: {},
    },
  });
  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 22,
    method: "tools/call",
    params: {
      name: "powerbi_diagnostics",
      arguments: {},
    },
  });

  await new Promise((resolve) => setImmediate(resolve));

  const response = asObject(written.find((payload) => isJsonObject(payload) && payload.id === 22));
  const diagnostics = diagnosticsFromResult(asObject(response.result));
  const upstream = asObject(diagnostics.upstream);
  assert.equal(upstream.status, "failed");
  assert.equal(upstream.error, "Remote MCP HTTP 503");
  assert.doesNotMatch(JSON.stringify(response), /user@example\.com|msal_token_cache|cached-token|Bearer|sensitive/i);
});

test("local Power BI wrapper tools invoke upstream Fabric MCP tools", async () => {
  const written: JSONValue[] = [];
  const forwarded: JSONRPCPayload[] = [];
  const localTools = new LocalPowerBITools({
    async getCachedAccessToken(): Promise<string> {
      return "cached-token";
    },
  } as never);

  const transport = new MCPProxyTransport(
    {
      async forward(payload: JSONRPCPayload): Promise<JSONValue> {
        forwarded.push(payload);
        if (!Array.isArray(payload) && payload.method === "initialize") {
          return {
            jsonrpc: "2.0",
            id: normalizeId(payload.id),
            result: {
              capabilities: {
                tools: {},
              },
            },
          };
        }
        if (!Array.isArray(payload) && payload.method === "tools/list") {
          return {
            jsonrpc: "2.0",
            id: normalizeId(payload.id),
            result: {
              tools: [
                {
                  name: "fabric_list_workspaces",
                  description: "List available Power BI workspaces.",
                  inputSchema: {
                    type: "object",
                    properties: {},
                  },
                },
              ],
            },
          };
        }
        if (!Array.isArray(payload) && payload.method === "tools/call") {
          const params = asObject(payload.params);
          return {
            jsonrpc: "2.0",
            id: normalizeId(payload.id),
            result: {
              content: [
                {
                  type: "text",
                  text: `called ${params.name}`,
                },
              ],
            },
          };
        }
        return null;
      },
    },
    (payload) => written.push(payload),
    localTools,
  );

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 0,
    method: "initialize",
    params: {
      protocolVersion: "2025-11-25",
      capabilities: {},
    },
  });

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 3,
    method: "tools/call",
    params: {
      name: "powerbi_list_workspaces",
      arguments: {},
    },
  });

  await new Promise((resolve) => setImmediate(resolve));

  const remoteCall = forwarded.find((payload) => !Array.isArray(payload) && payload.method === "tools/call");
  assert.ok(remoteCall && !Array.isArray(remoteCall));
  assert.equal(asObject(remoteCall.params).name, "fabric_list_workspaces");

  const response = asObject(written.find((payload) => isJsonObject(payload) && payload.id === 3));
  const result = asObject(response.result);
  const content = result.content;
  assert.ok(Array.isArray(content));
  assert.equal(asObject(content[0]).text, "called fabric_list_workspaces");
});

test("local list wrappers fall back to configured default context when Fabric MCP has no list tool", async () => {
  const written: JSONValue[] = [];
  const localTools = new LocalPowerBITools(
    {
      async getCachedAccessToken(): Promise<string> {
        return "cached-token";
      },
    } as never,
    {
      defaultWorkspaceId: "workspace-123",
      defaultWorkspaceName: "Finance",
      defaultSemanticModelId: "model-456",
      defaultSemanticModelName: "Finance Model",
    },
  );

  const transport = new MCPProxyTransport(
    {
      async forward(payload: JSONRPCPayload): Promise<JSONValue> {
        if (!Array.isArray(payload) && payload.method === "initialize") {
          return {
            jsonrpc: "2.0",
            id: normalizeId(payload.id),
            result: {
              capabilities: {
                tools: {},
              },
            },
          };
        }
        if (!Array.isArray(payload) && payload.method === "tools/list") {
          return {
            jsonrpc: "2.0",
            id: normalizeId(payload.id),
            result: {
              tools: [
                {
                  name: "GetSemanticModelSchema",
                  description: "Gets schema for a semantic model.",
                  inputSchema: {
                    type: "object",
                    properties: {
                      artifactId: {
                        type: "string",
                      },
                    },
                    required: ["artifactId"],
                  },
                },
              ],
            },
          };
        }
        return null;
      },
    },
    (payload) => written.push(payload),
    localTools,
  );

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 0,
    method: "initialize",
    params: {
      protocolVersion: "2025-11-25",
      capabilities: {},
    },
  });
  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 4,
    method: "tools/call",
    params: {
      name: "powerbi_list_semantic_models",
      arguments: {},
    },
  });

  await new Promise((resolve) => setImmediate(resolve));

  const response = asObject(written.find((payload) => isJsonObject(payload) && payload.id === 4));
  const result = asObject(response.result);
  const content = result.content;
  assert.ok(Array.isArray(content));
  assert.match(String(asObject(content[0]).text), /Finance Model/);
  assert.match(String(asObject(content[0]).text), /model-456/);
});

test("direct upstream tool calls receive configured default semantic model IDs", async () => {
  const written: JSONValue[] = [];
  const forwarded: JSONRPCPayload[] = [];
  const localTools = new LocalPowerBITools(
    {
      async getCachedAccessToken(): Promise<string> {
        return "cached-token";
      },
    } as never,
    {
      defaultSemanticModelId: "default-model-id",
    },
  );

  const transport = new MCPProxyTransport(
    {
      async forward(payload: JSONRPCPayload): Promise<JSONValue> {
        forwarded.push(payload);
        if (!Array.isArray(payload) && payload.method === "initialize") {
          return {
            jsonrpc: "2.0",
            id: normalizeId(payload.id),
            result: {
              capabilities: {
                tools: {},
              },
            },
          };
        }
        if (!Array.isArray(payload) && payload.method === "tools/list") {
          return semanticSchemaTools(payload);
        }
        if (!Array.isArray(payload) && payload.method === "tools/call") {
          return {
            jsonrpc: "2.0",
            id: normalizeId(payload.id),
            result: {
              content: [
                {
                  type: "text",
                  text: "ok",
                },
              ],
            },
          };
        }
        return null;
      },
    },
    (payload) => written.push(payload),
    localTools,
  );

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 0,
    method: "initialize",
    params: {
      protocolVersion: "2025-11-25",
      capabilities: {},
    },
  });
  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 5,
    method: "tools/call",
    params: {
      name: "GetSemanticModelSchema",
      arguments: {},
    },
  });

  await new Promise((resolve) => setImmediate(resolve));

  const remoteCall = forwarded.find((payload) => !Array.isArray(payload) && payload.method === "tools/call");
  assert.ok(remoteCall && !Array.isArray(remoteCall));
  const remoteArgs = asObject(asObject(remoteCall.params).arguments);
  assert.equal(remoteArgs.artifactId, "default-model-id");
});

test("local query wrappers call upstream Fabric query tools with active semantic model context", async () => {
  const forwarded: JSONRPCPayload[] = [];
  const localTools = new LocalPowerBITools(
    {
      async getCachedAccessToken(): Promise<string> {
        return "cached-token";
      },
    } as never,
    {
      defaultSemanticModelId: "default-model-id",
    },
  );

  const transport = new MCPProxyTransport(
    {
      async forward(payload: JSONRPCPayload): Promise<JSONValue> {
        forwarded.push(payload);
        if (!Array.isArray(payload) && payload.method === "initialize") {
          return {
            jsonrpc: "2.0",
            id: normalizeId(payload.id),
            result: {
              capabilities: {
                tools: {},
              },
            },
          };
        }
        if (!Array.isArray(payload) && payload.method === "tools/list") {
          return queryTools(payload);
        }
        if (!Array.isArray(payload) && payload.method === "tools/call") {
          return {
            jsonrpc: "2.0",
            id: normalizeId(payload.id),
            result: {
              content: [
                {
                  type: "text",
                  text: "ok",
                },
              ],
            },
          };
        }
        return null;
      },
    },
    () => undefined,
    localTools,
  );

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 0,
    method: "initialize",
    params: {
      protocolVersion: "2025-11-25",
      capabilities: {},
    },
  });
  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 15,
    method: "tools/call",
    params: {
      name: "powerbi_execute_dax_query",
      arguments: {
        daxQuery: "EVALUATE ROW(\"x\", 1)",
        maxRows: 10,
      },
    },
  });

  await new Promise((resolve) => setImmediate(resolve));

  const remoteCall = forwarded.find((payload) => !Array.isArray(payload) && payload.method === "tools/call");
  assert.ok(remoteCall && !Array.isArray(remoteCall));
  const params = asObject(remoteCall.params);
  const remoteArgs = asObject(params.arguments);
  assert.equal(params.name, "ExecuteQuery");
  assert.equal(remoteArgs.artifactId, "default-model-id");
  assert.deepEqual(remoteArgs.daxQueries, ["EVALUATE ROW(\"x\", 1)"]);
  assert.equal(remoteArgs.maxRows, 10);
});

test("local schema and generate wrappers map to upstream Fabric MCP tools", async () => {
  const forwarded: JSONRPCPayload[] = [];
  const localTools = new LocalPowerBITools(
    {
      async getCachedAccessToken(): Promise<string> {
        return "cached-token";
      },
    } as never,
    {
      defaultSemanticModelId: "default-model-id",
    },
  );

  const transport = new MCPProxyTransport(
    {
      async forward(payload: JSONRPCPayload): Promise<JSONValue> {
        forwarded.push(payload);
        if (!Array.isArray(payload) && payload.method === "initialize") {
          return {
            jsonrpc: "2.0",
            id: normalizeId(payload.id),
            result: {
              capabilities: {
                tools: {},
              },
            },
          };
        }
        if (!Array.isArray(payload) && payload.method === "tools/list") {
          return queryTools(payload);
        }
        if (!Array.isArray(payload) && payload.method === "tools/call") {
          return {
            jsonrpc: "2.0",
            id: normalizeId(payload.id),
            result: {
              content: [],
            },
          };
        }
        return null;
      },
    },
    () => undefined,
    localTools,
  );

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 0,
    method: "initialize",
    params: {
      protocolVersion: "2025-11-25",
      capabilities: {},
    },
  });
  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 16,
    method: "tools/call",
    params: {
      name: "powerbi_get_semantic_model_schema",
      arguments: {},
    },
  });
  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 17,
    method: "tools/call",
    params: {
      name: "powerbi_generate_dax_query",
      arguments: {
        question: "Show monthly revenue",
        schemaSelection: {
          tables: [
            {
              name: "Revenue",
              columns: ["Month"],
              measures: ["Revenue"],
            },
          ],
        },
      },
    },
  });

  await new Promise((resolve) => setImmediate(resolve));

  const calls = forwarded.filter((payload) => !Array.isArray(payload) && payload.method === "tools/call");
  const schemaCall = calls.find((payload) => !Array.isArray(payload) && asObject(payload.params).name === "GetSemanticModelSchema");
  const generateCall = calls.find((payload) => !Array.isArray(payload) && asObject(payload.params).name === "GenerateQuery");
  assert.ok(schemaCall && !Array.isArray(schemaCall));
  assert.ok(generateCall && !Array.isArray(generateCall));
  assert.equal(asObject(asObject(schemaCall.params).arguments).artifactId, "default-model-id");
  const generateArgs = asObject(asObject(generateCall.params).arguments);
  assert.equal(generateArgs.artifactId, "default-model-id");
  assert.equal(generateArgs.userInput, "Show monthly revenue");
  assert.ok(isJsonObject(generateArgs.schemaSelection));
});

test("chat context overrides configured defaults for upstream tool calls", async () => {
  const written: JSONValue[] = [];
  const forwarded: JSONRPCPayload[] = [];
  const localTools = new LocalPowerBITools(
    {
      async getCachedAccessToken(): Promise<string> {
        return "cached-token";
      },
    } as never,
    {
      defaultSemanticModelId: "default-model-id",
      defaultSemanticModelName: "Default Model",
    },
  );

  const transport = new MCPProxyTransport(
    {
      async forward(payload: JSONRPCPayload): Promise<JSONValue> {
        forwarded.push(payload);
        if (!Array.isArray(payload) && payload.method === "initialize") {
          return {
            jsonrpc: "2.0",
            id: normalizeId(payload.id),
            result: {
              capabilities: {
                tools: {},
              },
            },
          };
        }
        if (!Array.isArray(payload) && payload.method === "tools/list") {
          return semanticSchemaTools(payload);
        }
        if (!Array.isArray(payload) && payload.method === "tools/call") {
          return {
            jsonrpc: "2.0",
            id: normalizeId(payload.id),
            result: {
              content: [
                {
                  type: "text",
                  text: "ok",
                },
              ],
            },
          };
        }
        return null;
      },
    },
    (payload) => written.push(payload),
    localTools,
  );

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 0,
    method: "initialize",
    params: {
      protocolVersion: "2025-11-25",
      capabilities: {},
    },
  });
  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 8,
    method: "tools/call",
    params: {
      name: "powerbi_set_context",
      arguments: {
        semanticModelId: "chat-model-id",
        semanticModelName: "Chat Model",
      },
    },
  });

  await new Promise((resolve) => setImmediate(resolve));

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 9,
    method: "tools/call",
    params: {
      name: "GetSemanticModelSchema",
      arguments: {},
    },
  });

  await new Promise((resolve) => setImmediate(resolve));

  const setResponse = asObject(written.find((payload) => isJsonObject(payload) && payload.id === 8));
  const setResult = asObject(setResponse.result);
  const setContent = setResult.content;
  assert.ok(Array.isArray(setContent));
  assert.match(String(asObject(setContent[0]).text), /Chat Model/);

  const remoteCall = forwarded.find((payload) => !Array.isArray(payload) && payload.method === "tools/call");
  assert.ok(remoteCall && !Array.isArray(remoteCall));
  const remoteArgs = asObject(asObject(remoteCall.params).arguments);
  assert.equal(remoteArgs.artifactId, "chat-model-id");
});

test("explicit semantic model IDs win over chat context", async () => {
  const forwarded: JSONRPCPayload[] = [];
  const localTools = new LocalPowerBITools(
    {
      async getCachedAccessToken(): Promise<string> {
        return "cached-token";
      },
    } as never,
    {
      defaultSemanticModelId: "default-model-id",
    },
  );

  const transport = new MCPProxyTransport(
    {
      async forward(payload: JSONRPCPayload): Promise<JSONValue> {
        forwarded.push(payload);
        if (!Array.isArray(payload) && payload.method === "initialize") {
          return {
            jsonrpc: "2.0",
            id: normalizeId(payload.id),
            result: {
              capabilities: {
                tools: {},
              },
            },
          };
        }
        if (!Array.isArray(payload) && payload.method === "tools/list") {
          return semanticSchemaTools(payload);
        }
        if (!Array.isArray(payload) && payload.method === "tools/call") {
          return {
            jsonrpc: "2.0",
            id: normalizeId(payload.id),
            result: {
              content: [],
            },
          };
        }
        return null;
      },
    },
    () => undefined,
    localTools,
  );

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 0,
    method: "initialize",
    params: {
      protocolVersion: "2025-11-25",
      capabilities: {},
    },
  });
  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 10,
    method: "tools/call",
    params: {
      name: "powerbi_set_context",
      arguments: {
        semanticModelId: "chat-model-id",
      },
    },
  });

  await new Promise((resolve) => setImmediate(resolve));

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 11,
    method: "tools/call",
    params: {
      name: "GetSemanticModelSchema",
      arguments: {
        artifactId: "explicit-model-id",
      },
    },
  });

  await new Promise((resolve) => setImmediate(resolve));

  const remoteCall = forwarded.find((payload) => !Array.isArray(payload) && payload.method === "tools/call");
  assert.ok(remoteCall && !Array.isArray(remoteCall));
  const remoteArgs = asObject(asObject(remoteCall.params).arguments);
  assert.equal(remoteArgs.artifactId, "explicit-model-id");
});

test("clearing chat context restores configured defaults", async () => {
  const forwarded: JSONRPCPayload[] = [];
  const localTools = new LocalPowerBITools(
    {
      async getCachedAccessToken(): Promise<string> {
        return "cached-token";
      },
    } as never,
    {
      defaultSemanticModelId: "default-model-id",
    },
  );

  const transport = new MCPProxyTransport(
    {
      async forward(payload: JSONRPCPayload): Promise<JSONValue> {
        forwarded.push(payload);
        if (!Array.isArray(payload) && payload.method === "initialize") {
          return {
            jsonrpc: "2.0",
            id: normalizeId(payload.id),
            result: {
              capabilities: {
                tools: {},
              },
            },
          };
        }
        if (!Array.isArray(payload) && payload.method === "tools/list") {
          return semanticSchemaTools(payload);
        }
        if (!Array.isArray(payload) && payload.method === "tools/call") {
          return {
            jsonrpc: "2.0",
            id: normalizeId(payload.id),
            result: {
              content: [],
            },
          };
        }
        return null;
      },
    },
    () => undefined,
    localTools,
  );

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 0,
    method: "initialize",
    params: {
      protocolVersion: "2025-11-25",
      capabilities: {},
    },
  });
  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 12,
    method: "tools/call",
    params: {
      name: "powerbi_set_context",
      arguments: {
        semanticModelId: "chat-model-id",
      },
    },
  });
  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 13,
    method: "tools/call",
    params: {
      name: "powerbi_clear_context",
      arguments: {},
    },
  });

  await new Promise((resolve) => setImmediate(resolve));

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 14,
    method: "tools/call",
    params: {
      name: "GetSemanticModelSchema",
      arguments: {},
    },
  });

  await new Promise((resolve) => setImmediate(resolve));

  const remoteCall = forwarded.find((payload) => !Array.isArray(payload) && payload.method === "tools/call");
  assert.ok(remoteCall && !Array.isArray(remoteCall));
  const remoteArgs = asObject(asObject(remoteCall.params).arguments);
  assert.equal(remoteArgs.artifactId, "default-model-id");
});

test("direct upstream tool calls preserve explicit semantic model IDs", async () => {
  const forwarded: JSONRPCPayload[] = [];
  const localTools = new LocalPowerBITools(
    {
      async getCachedAccessToken(): Promise<string> {
        return "cached-token";
      },
    } as never,
    {
      defaultSemanticModelId: "default-model-id",
    },
  );

  const transport = new MCPProxyTransport(
    {
      async forward(payload: JSONRPCPayload): Promise<JSONValue> {
        forwarded.push(payload);
        if (!Array.isArray(payload) && payload.method === "initialize") {
          return {
            jsonrpc: "2.0",
            id: normalizeId(payload.id),
            result: {
              capabilities: {
                tools: {},
              },
            },
          };
        }
        if (!Array.isArray(payload) && payload.method === "tools/list") {
          return semanticSchemaTools(payload);
        }
        if (!Array.isArray(payload) && payload.method === "tools/call") {
          return {
            jsonrpc: "2.0",
            id: normalizeId(payload.id),
            result: {
              content: [],
            },
          };
        }
        return null;
      },
    },
    () => undefined,
    localTools,
  );

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 0,
    method: "initialize",
    params: {
      protocolVersion: "2025-11-25",
      capabilities: {},
    },
  });
  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 6,
    method: "tools/call",
    params: {
      name: "GetSemanticModelSchema",
      arguments: {
        artifactId: "explicit-model-id",
      },
    },
  });

  await new Promise((resolve) => setImmediate(resolve));

  const remoteCall = forwarded.find((payload) => !Array.isArray(payload) && payload.method === "tools/call");
  assert.ok(remoteCall && !Array.isArray(remoteCall));
  const remoteArgs = asObject(asObject(remoteCall.params).arguments);
  assert.equal(remoteArgs.artifactId, "explicit-model-id");
});

test("direct upstream tool calls report missing required default IDs", async () => {
  const written: JSONValue[] = [];
  const forwarded: JSONRPCPayload[] = [];
  const localTools = new LocalPowerBITools({
    async getCachedAccessToken(): Promise<string> {
      return "cached-token";
    },
  } as never);

  const transport = new MCPProxyTransport(
    {
      async forward(payload: JSONRPCPayload): Promise<JSONValue> {
        forwarded.push(payload);
        if (!Array.isArray(payload) && payload.method === "initialize") {
          return {
            jsonrpc: "2.0",
            id: normalizeId(payload.id),
            result: {
              capabilities: {
                tools: {},
              },
            },
          };
        }
        if (!Array.isArray(payload) && payload.method === "tools/list") {
          return semanticSchemaTools(payload);
        }
        throw new Error("Should not forward missing required ID tools/call.");
      },
    },
    (payload) => written.push(payload),
    localTools,
  );

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 0,
    method: "initialize",
    params: {
      protocolVersion: "2025-11-25",
      capabilities: {},
    },
  });
  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 7,
    method: "tools/call",
    params: {
      name: "GetSemanticModelSchema",
      arguments: {},
    },
  });

  await new Promise((resolve) => setImmediate(resolve));

  assert.equal(
    forwarded.some((payload) => !Array.isArray(payload) && payload.method === "tools/call"),
    false,
  );
  const response = asObject(written.find((payload) => isJsonObject(payload) && payload.id === 7));
  const result = asObject(response.result);
  const content = result.content;
  assert.ok(Array.isArray(content));
  assert.match(String(asObject(content[0]).text), /default_semantic_model_id/);
});

test("default argument injection supports common Power BI ID field names", () => {
  const fields = ["workspaceId", "groupId", "semanticModelId", "datasetId", "artifactId"];
  const tool: RemoteTool = {
    name: "PowerBITool",
    inputSchema: {
      type: "object",
      properties: Object.fromEntries(fields.map((field) => [field, { type: "string" }])),
      required: fields,
    },
  };

  const result = applyDefaultArguments(tool, {}, {
    workspaceId: "default-workspace-id",
    semanticModelId: "default-model-id",
  });

  assert.equal(result.status, "ok");
  if (result.status !== "ok") {
    throw new Error("Expected default argument injection to succeed.");
  }
  assert.equal(result.arguments.workspaceId, "default-workspace-id");
  assert.equal(result.arguments.groupId, "default-workspace-id");
  assert.equal(result.arguments.semanticModelId, "default-model-id");
  assert.equal(result.arguments.datasetId, "default-model-id");
  assert.equal(result.arguments.artifactId, "default-model-id");
});

test("device-code prompt is unavailable without URL-mode elicitation", async () => {
  const written: JSONValue[] = [];
  const transport = new MCPProxyTransport(
    {
      async forward(): Promise<JSONValue> {
        return null;
      },
    },
    (payload) => written.push(payload),
  );

  transport.handleIncoming({
    jsonrpc: "2.0",
    id: 1,
    method: "initialize",
    params: {
      capabilities: {
        elicitation: {},
      },
    },
  });

  assert.equal(transport.supportsDeviceCodePrompt(), false);
  assert.equal(await transport.promptDeviceCode(deviceCode), "unavailable");
  assert.equal(
    written.some((payload) => typeof payload === "object" && payload !== null && !Array.isArray(payload) && payload.method === "elicitation/create"),
    false,
  );
});

function findElicitation(written: JSONValue[]): JSONObject {
  const payload = written.find(
    (candidate) => isJsonObject(candidate) && candidate.method === "elicitation/create",
  );
  if (!isJsonObject(payload)) {
    throw new Error("Expected elicitation/create payload.");
  }
  return payload;
}

function asObject(value: JSONValue | undefined): JSONObject {
  if (!isJsonObject(value)) {
    throw new Error("Expected JSON object.");
  }
  return value;
}

function isJsonObject(value: unknown): value is JSONObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function diagnosticsFromResult(result: JSONObject): JSONObject {
  const structuredContent = asObject(result.structuredContent);
  return asObject(structuredContent.diagnostics);
}

function semanticSchemaTools(payload: JSONObject): JSONObject {
  return {
    jsonrpc: "2.0",
    id: normalizeId(payload.id),
    result: {
      tools: [
        {
          name: "GetSemanticModelSchema",
          description: "Gets schema for a semantic model.",
          inputSchema: {
            type: "object",
            properties: {
              artifactId: {
                type: "string",
              },
            },
            required: ["artifactId"],
          },
        },
      ],
    },
  };
}

function queryTools(payload: JSONObject): JSONObject {
  return {
    jsonrpc: "2.0",
    id: normalizeId(payload.id),
    result: {
      tools: [
        {
          name: "ExecuteQuery",
          description: "Executes a DAX query and return the results.",
          inputSchema: {
            type: "object",
            properties: {
              artifactId: {
                type: "string",
              },
              daxQueries: {
                type: "array",
                items: {
                  type: "string",
                },
              },
              maxRows: {
                type: ["integer", "null"],
              },
            },
            required: ["artifactId", "daxQueries"],
          },
        },
        {
          name: "GetSemanticModelSchema",
          description: "Gets schema for a semantic model.",
          inputSchema: {
            type: "object",
            properties: {
              artifactId: {
                type: "string",
              },
            },
            required: ["artifactId"],
          },
        },
        {
          name: "GenerateQuery",
          description: "Generates a DAX query for a semantic model.",
          inputSchema: {
            type: "object",
            properties: {
              artifactId: {
                type: "string",
              },
              userInput: {
                type: "string",
              },
              schemaSelection: {
                type: "object",
              },
            },
            required: ["artifactId", "userInput"],
          },
        },
        {
          name: "GetReportMetadata",
          description: "Gets report metadata.",
          inputSchema: {
            type: "object",
            properties: {
              reportObjectId: {
                type: "string",
              },
            },
            required: ["reportObjectId"],
          },
        },
      ],
    },
  };
}

test("SSE payload parser returns one JSON-RPC event or a list of events", () => {
  assert.deepEqual(
    parseSsePayload('event: message\ndata: {"jsonrpc":"2.0","id":1,"result":{"ok":true}}\n\n'),
    {
      jsonrpc: "2.0",
      id: 1,
      result: {
        ok: true,
      },
    },
  );

  assert.deepEqual(
    parseSsePayload('data: {"method":"a"}\n\ndata: {"method":"b"}\n\n'),
    [
      {
        method: "a",
      },
      {
        method: "b",
      },
    ],
  );
});
