import assert from "node:assert/strict";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { AddressInfo } from "node:net";
import { test, type TestContext } from "node:test";

import type { ProxyConfig } from "../src/config.js";
import { type JSONObject, normalizeId, type JSONRPCPayload, type JSONValue } from "../src/jsonrpc.js";
import { createHttpMCPServer, sessionConfig, type HostedMCPServerOptions } from "../src/httpServer.js";
import { LocalPowerBITools } from "../src/localTools.js";
import type { RemoteForwarder } from "../src/remote.js";
import { MCPProxyTransport } from "../src/transport.js";

test("hosted HTTP server exposes health and initialize with a session id", async (t) => {
  const harness = await createHarness(t);

  const health = await fetch(`${harness.baseUrl}/healthz`);
  assert.equal(health.status, 200);
  assert.equal(asObject(await health.json()).ok, true);

  const initialized = await harness.post(initializePayload(1));
  assert.equal(initialized.status, 200);
  assert.ok(initialized.sessionId);

  const body = asObject(initialized.body);
  const result = asObject(body.result);
  const serverInfo = asObject(result.serverInfo);
  assert.equal(serverInfo.name, "powerbi-mcp-claude");
});

test("hosted tools/list works over HTTP", async (t) => {
  const harness = await createHarness(t);
  const initialized = await harness.post(initializePayload(1));

  const listed = await harness.post(
    {
      jsonrpc: "2.0",
      id: 2,
      method: "tools/list",
    },
    initialized.sessionId,
  );

  assert.equal(listed.status, 200);
  const tools = asObject(asObject(listed.body).result).tools;
  assert.ok(Array.isArray(tools));
  const toolNames = tools.map((tool) => asObject(tool).name);
  assert.ok(toolNames.includes("powerbi_auth_start"));
  assert.ok(toolNames.includes("powerbi_set_context"));
  assert.ok(toolNames.includes("powerbi_execute_dax_query"));
});

test("hosted sessions receive isolated MSAL token cache paths", async (t) => {
  const harness = await createHarness(t);
  const first = await harness.post(initializePayload(1));
  const second = await harness.post(initializePayload(2));

  assert.ok(first.sessionId);
  assert.ok(second.sessionId);
  assert.notEqual(first.sessionId, second.sessionId);

  const firstConfig = harness.createdConfigs.find((config) => config.cacheDir.endsWith(first.sessionId ?? ""));
  const secondConfig = harness.createdConfigs.find((config) => config.cacheDir.endsWith(second.sessionId ?? ""));
  assert.ok(firstConfig);
  assert.ok(secondConfig);
  assert.match(firstConfig.tokenCachePath, new RegExp(`sessions/${escapeRegExp(first.sessionId ?? "")}/msal_token_cache\\.json$`));
  assert.match(secondConfig.tokenCachePath, new RegExp(`sessions/${escapeRegExp(second.sessionId ?? "")}/msal_token_cache\\.json$`));
  assert.notEqual(firstConfig.tokenCachePath, secondConfig.tokenCachePath);
});

test("powerbi_auth_start in one hosted session does not authenticate another session", async (t) => {
  const harness = await createHarness(t);
  const first = await harness.post(initializePayload(1));
  const second = await harness.post(initializePayload(2));

  const started = await harness.post(authToolPayload(3, "powerbi_auth_start"), first.sessionId);
  assert.match(resultText(started.body), /Microsoft sign-in is required/);

  const firstStatus = await harness.post(authToolPayload(4, "powerbi_auth_status"), first.sessionId);
  const secondStatus = await harness.post(authToolPayload(5, "powerbi_auth_status"), second.sessionId);
  assert.match(resultText(firstStatus.body), /Still waiting for Microsoft sign-in/);
  assert.match(resultText(secondStatus.body), /Microsoft authentication has not been started/);
});

test("powerbi_auth_logout clears only the current hosted session", async (t) => {
  const harness = await createHarness(t);
  const first = await harness.post(initializePayload(1));
  const second = await harness.post(initializePayload(2));

  await harness.post(authToolPayload(3, "powerbi_auth_start"), first.sessionId);
  await harness.post(authToolPayload(4, "powerbi_auth_start"), second.sessionId);

  const logout = await harness.post(authToolPayload(5, "powerbi_auth_logout"), first.sessionId);
  assert.match(resultText(logout.body), /Signed out of Power BI locally/);

  const firstStatus = await harness.post(authToolPayload(6, "powerbi_auth_status"), first.sessionId);
  const secondStatus = await harness.post(authToolPayload(7, "powerbi_auth_status"), second.sessionId);
  assert.match(resultText(firstStatus.body), /Microsoft authentication has not been started/);
  assert.match(resultText(secondStatus.body), /Still waiting for Microsoft sign-in/);
});

test("hosted chat context is isolated per MCP session", async (t) => {
  const harness = await createHarness(t);
  const first = await harness.post(initializePayload(1));
  const second = await harness.post(initializePayload(2));

  await harness.post(
    {
      jsonrpc: "2.0",
      id: 3,
      method: "tools/call",
      params: {
        name: "powerbi_set_context",
        arguments: {
          semanticModelId: "session-one-model",
          semanticModelName: "Session One Model",
        },
      },
    },
    first.sessionId,
  );

  const firstContext = await harness.post(authToolPayload(4, "powerbi_get_default_context"), first.sessionId);
  const secondContext = await harness.post(authToolPayload(5, "powerbi_get_default_context"), second.sessionId);

  assert.match(resultText(firstContext.body), /Session One Model/);
  assert.doesNotMatch(resultText(secondContext.body), /Session One Model|session-one-model/);
});

test("expired hosted sessions are replaced and do not reuse stale chat context", async (t) => {
  let now = 0;
  const harness = await createHarness(t, {
    sessionTtlMs: 1_000,
    now: () => now,
  });
  const first = await harness.post(initializePayload(1));
  assert.ok(first.sessionId);

  await harness.post(
    {
      jsonrpc: "2.0",
      id: 2,
      method: "tools/call",
      params: {
        name: "powerbi_set_context",
        arguments: {
          semanticModelId: "expired-model",
          semanticModelName: "Expired Model",
        },
      },
    },
    first.sessionId,
  );

  now = 2_000;
  const afterExpiry = await harness.post(authToolPayload(3, "powerbi_get_default_context"), first.sessionId);
  assert.ok(afterExpiry.sessionId);
  assert.notEqual(afterExpiry.sessionId, first.sessionId);
  assert.doesNotMatch(resultText(afterExpiry.body), /Expired Model|expired-model/);
});

test("sessionConfig scopes hosted token cache under sessions directory", () => {
  const base = testConfig("/tmp/powerbi-mcp-test");
  const scoped = sessionConfig(base, "session-id");
  assert.equal(scoped.cacheDir, "/tmp/powerbi-mcp-test/sessions/session-id");
  assert.equal(scoped.tokenCachePath, "/tmp/powerbi-mcp-test/sessions/session-id/msal_token_cache.json");
  assert.equal(base.tokenCachePath, "/tmp/powerbi-mcp-test/msal_token_cache.json");
});

type HarnessOptions = {
  sessionTtlMs?: number;
  now?: () => number;
};

type PostResult = {
  status: number;
  sessionId: string | null;
  body: JSONValue | undefined;
};

async function createHarness(
  t: TestContext,
  options: HarnessOptions = {},
): Promise<{
  baseUrl: string;
  createdConfigs: ProxyConfig[];
  post: (payload: JSONValue, sessionId?: string | null) => Promise<PostResult>;
}> {
  const tempDir = mkdtempSync(join(tmpdir(), "powerbi-mcp-http-test-"));
  const config = testConfig(tempDir);
  const createdConfigs: ProxyConfig[] = [];
  const serverOptions: HostedMCPServerOptions = {
    config,
    transportFactory: (sessionScopedConfig) => {
      createdConfigs.push(sessionScopedConfig);
      return new MCPProxyTransport(
        fakeRemote,
        () => undefined,
        new LocalPowerBITools(fakeAuth(sessionScopedConfig) as never, sessionScopedConfig),
      );
    },
  };
  if (options.sessionTtlMs !== undefined) {
    serverOptions.sessionTtlMs = options.sessionTtlMs;
  }
  if (options.now !== undefined) {
    serverOptions.now = options.now;
  }
  const server = createHttpMCPServer(serverOptions);

  await new Promise<void>((resolve) => server.listen(0, "127.0.0.1", resolve));
  t.after(() => {
    server.close();
    rmSync(tempDir, { recursive: true, force: true });
  });

  const address = server.address() as AddressInfo;
  const baseUrl = `http://127.0.0.1:${address.port}`;

  return {
    baseUrl,
    createdConfigs,
    post: async (payload, sessionId) => {
      const headers: Record<string, string> = {
        "Content-Type": "application/json",
      };
      if (sessionId) {
        headers["Mcp-Session-Id"] = sessionId;
      }

      const response = await fetch(`${baseUrl}/mcp`, {
        method: "POST",
        headers,
        body: JSON.stringify(payload),
      });
      const text = await response.text();
      return {
        status: response.status,
        sessionId: response.headers.get("mcp-session-id"),
        body: text ? (JSON.parse(text) as JSONValue) : undefined,
      };
    },
  };
}

function testConfig(cacheDir: string): ProxyConfig {
  return {
    remoteUrl: "https://api.fabric.microsoft.com/v1/mcp/powerbi",
    clientId: "test-client-id",
    tenantId: "organizations",
    authorityHost: "https://login.microsoftonline.com",
    scopes: ["https://analysis.windows.net/powerbi/api/.default"],
    cacheDir,
    tokenCachePath: join(cacheDir, "msal_token_cache.json"),
    timeoutSeconds: 60,
    clientName: "powerbi-mcp-test",
    clientVersion: "0.0.0-test",
    allowInteractiveAuth: true,
    deviceCodeTimeoutSeconds: 900,
  };
}

function fakeAuth(config: ProxyConfig): object {
  let started = false;
  return {
    async getCachedAccessToken(): Promise<undefined> {
      return undefined;
    },
    async startDeviceCodeLogin() {
      started = true;
      return {
        status: "pending",
        userCode: "TEST-CODE",
        verificationUri: "https://login.microsoft.com/device",
        expiresIn: 900,
        message: "Open https://login.microsoft.com/device and enter code TEST-CODE to authenticate with Microsoft.",
      };
    },
    async deviceLoginStatus() {
      if (!started) {
        return {
          status: "not_started",
          authenticated: false,
          message: "Microsoft authentication has not been started.",
        };
      }

      return {
        status: "pending",
        authenticated: false,
        userCode: "TEST-CODE",
        verificationUri: "https://login.microsoft.com/device",
        message: "Still waiting for Microsoft sign-in. Open https://login.microsoft.com/device and enter code TEST-CODE.",
      };
    },
    async clearCache() {
      const accountsRemoved = started ? 1 : 0;
      started = false;
      return {
        cachePath: config.tokenCachePath,
        cacheExists: false,
        accountsRemoved,
        remainingAccounts: 0,
      };
    },
    async tokenInfo() {
      return {
        cachePath: config.tokenCachePath,
        cacheExists: started,
        accountCount: started ? 1 : 0,
        usernames: [],
      };
    },
  };
}

const fakeRemote: RemoteForwarder = {
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
          tools: [],
        },
      };
    }

    return null;
  },
};

function initializePayload(id: number): JSONObject {
  return {
    jsonrpc: "2.0",
    id,
    method: "initialize",
    params: {
      protocolVersion: "2025-11-25",
      capabilities: {},
    },
  };
}

function authToolPayload(id: number, name: string): JSONObject {
  return {
    jsonrpc: "2.0",
    id,
    method: "tools/call",
    params: {
      name,
      arguments: {},
    },
  };
}

function resultText(value: JSONValue | undefined): string {
  const result = asObject(asObject(value).result);
  const content = result.content;
  if (!Array.isArray(content)) {
    throw new Error("Expected tool content.");
  }
  return String(asObject(content[0]).text);
}

function asObject(value: JSONValue | undefined): JSONObject {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("Expected JSON object.");
  }
  return value;
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
