import assert from "node:assert/strict";
import { test } from "node:test";
import type { DeviceCodeResponse } from "@azure/msal-common/node";

import { normalizeId, type JSONObject, type JSONRPCPayload, type JSONValue } from "../src/jsonrpc.js";
import { LocalPowerBITools } from "../src/localTools.js";
import type { RemoteForwarder } from "../src/remote.js";
import { parseSsePayload } from "../src/remote.js";
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
  assert.ok(toolNames.includes("powerbi_list_workspaces"));
  assert.ok(toolNames.includes("powerbi_list_semantic_models"));
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
