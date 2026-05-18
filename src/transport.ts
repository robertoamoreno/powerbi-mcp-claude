import { createInterface } from "node:readline";
import type { Readable, Writable } from "node:stream";
import type { DeviceCodeResponse } from "@azure/msal-common/node";

import type { DeviceCodePrompter, DeviceCodePromptResult } from "./auth.js";
import {
  firstRequestId,
  idKey,
  INTERNAL_ERROR,
  INVALID_REQUEST,
  isJsonRpcPayload,
  isObject,
  isResponse,
  jsonRpcError,
  normalizeId,
  type JSONObject,
  type JSONRPCId,
  type JSONRPCPayload,
  type JSONValue,
} from "./jsonrpc.js";
import {
  AUTH_TOOL_NAMES,
  CONFIG_TOOL_NAMES,
  isLocalToolCall,
  mergeTools,
  POWERBI_TOOL_NAMES,
  toolsListResult,
  type LocalPowerBITools,
} from "./localTools.js";
import { RemoteMCPError, type RemoteForwarder } from "./remote.js";
import {
  applyDefaultArguments,
  extractRemoteTools,
  invokeRemotePowerBITool,
  signInRequiredResult,
  type RemoteTool,
} from "./remotePowerBITools.js";
import { log } from "./logger.js";
import { VERSION } from "./version.js";

type PendingRequest = {
  resolve: (value: JSONValue) => void;
  reject: (error: Error) => void;
  timer: NodeJS.Timeout;
};

type RemotePayloadPreparation =
  | {
      payload: JSONRPCPayload;
    }
  | {
      response: JSONValue;
    };

const DEFAULT_CLIENT_REQUEST_TIMEOUT_MS = 15 * 60 * 1000;
const EXPECTED_FABRIC_QUERY_TOOLS = [
  "ExecuteQuery",
  "GetSemanticModelSchema",
  "GenerateQuery",
  "GetReportMetadata",
] as const;
const EXPECTED_LOCAL_WRAPPER_TOOLS = [
  "powerbi_auth_start",
  "powerbi_auth_status",
  "powerbi_auth_logout",
  "powerbi_diagnostics",
  "powerbi_get_default_context",
  "powerbi_set_context",
  "powerbi_clear_context",
  "powerbi_list_workspaces",
  "powerbi_list_semantic_models",
  "powerbi_get_semantic_model_schema",
  "powerbi_generate_dax_query",
  "powerbi_execute_dax_query",
  "powerbi_get_report_metadata",
] as const;

export class MCPProxyTransport implements DeviceCodePrompter {
  private nextServerRequestId = 1;
  private readonly pending = new Map<string, PendingRequest>();
  private clientCapabilities: Record<string, unknown> = {};
  private initializePayload: JSONRPCPayload | undefined;
  private initializedNotification: JSONRPCPayload | undefined;
  private remoteInitialized = false;
  private remoteTools = new Map<string, RemoteTool>();

  constructor(
    private readonly remote: RemoteForwarder,
    private readonly writePayload: (payload: JSONValue) => void,
    private readonly localTools: LocalPowerBITools | undefined = undefined,
  ) {}

  async run(input: Readable): Promise<void> {
    log("Power BI MCP stdio proxy started.");
    const lines = createInterface({ input, crlfDelay: Infinity });

    for await (const raw of lines) {
      const line = raw.trim();
      if (!line) {
        continue;
      }

      let parsed: unknown;
      try {
        parsed = JSON.parse(line);
      } catch (error) {
        this.writePayload(jsonRpcError(null, -32700, `Parse error: ${errorMessage(error)}`));
        continue;
      }

      this.handleIncoming(parsed);
    }

    log("stdin closed; stopping proxy.");
  }

  handleIncoming(parsed: unknown): void {
    if (!isJsonRpcPayload(parsed)) {
      this.writePayload(jsonRpcError(null, INVALID_REQUEST, "Invalid JSON-RPC message."));
      return;
    }

    const payload = this.resolveInternalResponses(parsed);
    if (!payload) {
      return;
    }

    if (this.handleLocalInitialize(payload)) {
      return;
    }
    if (this.handleInitializedNotification(payload)) {
      return;
    }

    void this.forwardClientPayload(payload);
  }

  supportsDeviceCodePrompt(): boolean {
    const elicitation = this.clientCapabilities.elicitation;
    if (!isObject(elicitation)) {
      return false;
    }

    return isObject(elicitation.url);
  }

  async promptDeviceCode(response: DeviceCodeResponse): Promise<DeviceCodePromptResult> {
    if (!this.supportsDeviceCodePrompt()) {
      return "unavailable";
    }

    const elicitationId = `powerbi-device-code-${Date.now()}-${this.nextServerRequestId}`;
    const result = await this.sendClientRequest(
      "elicitation/create",
      {
        mode: "url",
        elicitationId,
        url: response.verificationUri,
        message: `Microsoft sign-in is required for Power BI. Open the Microsoft page and enter code ${response.userCode}.`,
      },
      Math.max(response.expiresIn * 1000, DEFAULT_CLIENT_REQUEST_TIMEOUT_MS),
    );

    if (!isObject(result)) {
      return "cancelled";
    }

    const action = result.action;
    if (action === "accept") {
      return "accepted";
    }
    if (action === "decline") {
      return "declined";
    }
    if (action === "cancel") {
      return "cancelled";
    }

    return "cancelled";
  }

  private async sendClientRequest(method: string, params: JSONObject, timeoutMs: number): Promise<JSONValue> {
    const id = `powerbi-proxy-${this.nextServerRequestId++}`;
    const key = idKey(id);

    const result = new Promise<JSONValue>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(key);
        reject(new Error(`Timed out waiting for client response to ${method}.`));
      }, timeoutMs);

      this.pending.set(key, { resolve, reject, timer });
    });

    this.writePayload({ jsonrpc: "2.0", id, method, params });
    return await result;
  }

  private resolveInternalResponses(payload: JSONRPCPayload): JSONRPCPayload | undefined {
    if (!Array.isArray(payload)) {
      return this.tryResolveInternalResponse(payload) ? undefined : payload;
    }

    const passthrough: JSONValue[] = [];
    for (const item of payload) {
      if (this.tryResolveInternalResponse(item)) {
        continue;
      }
      passthrough.push(item);
    }

    if (passthrough.length === 0) {
      return undefined;
    }

    return passthrough;
  }

  private tryResolveInternalResponse(value: unknown): boolean {
    if (!isResponse(value)) {
      return false;
    }

    const pending = this.pending.get(idKey(normalizeId(value.id)));
    if (!pending) {
      return false;
    }

    this.pending.delete(idKey(normalizeId(value.id)));
    clearTimeout(pending.timer);

    if ("error" in value && value.error) {
      pending.reject(new Error(formatJsonRpcError(value.error)));
    } else {
      pending.resolve(value.result ?? null);
    }
    return true;
  }

  private async forwardClientPayload(payload: JSONRPCPayload): Promise<void> {
    try {
      const localResponse = await this.handleLocalPayload(payload);
      if (localResponse !== undefined) {
        this.writePayload(localResponse);
        return;
      }

      await this.ensureRemoteInitialized();
      const prepared = await this.prepareRemotePayload(payload);
      if ("response" in prepared) {
        this.writePayload(prepared.response);
        return;
      }

      const response = await this.remote.forward(prepared.payload);
      if (response !== null) {
        this.writeServerResponse(response);
      }
    } catch (error) {
      log(errorMessage(error));
      this.writeForwardError(payload, error);
    }
  }

  private async handleLocalPayload(payload: JSONRPCPayload): Promise<JSONValue | undefined> {
    if (Array.isArray(payload) || !this.localTools) {
      return undefined;
    }

    if (payload.method === "tools/list") {
      const localTools = this.localTools.tools();
      const requestId = normalizeId(payload.id);
      if (!(await this.localTools.isAuthenticated())) {
        return toolsListResult(requestId, localTools);
      }

      try {
        await this.ensureRemoteInitialized();
        const remoteResponse = await this.remote.forward(payload);
        this.captureRemoteTools(remoteResponse, requestId);
        return remoteResponse === null ? toolsListResult(requestId, localTools) : mergeTools(remoteResponse, localTools);
      } catch (error) {
        log(`Could not load upstream Power BI MCP tools; exposing local tools only: ${errorMessage(error)}`);
        return toolsListResult(requestId, localTools);
      }
    }

    if (payload.method === "tools/call" && isObject(payload.params)) {
      const localToolName = isLocalToolCall(payload.params);
      if (localToolName) {
        const args = isObject(payload.params.arguments) ? (payload.params.arguments as JSONObject) : {};
        const result = await this.handleLocalToolCall(localToolName, args);
        return {
          jsonrpc: "2.0",
          id: normalizeId(payload.id),
          result,
        };
      }
    }

    return undefined;
  }

  private async handleLocalToolCall(name: string, args: JSONObject): Promise<JSONValue> {
    if (!this.localTools) {
      throw new Error("Local Power BI tools are not configured.");
    }

    if (name === "powerbi_diagnostics") {
      return await this.powerBIDiagnostics();
    }

    if (AUTH_TOOL_NAMES.has(name) || CONFIG_TOOL_NAMES.has(name)) {
      return await this.localTools.handleToolCall(name, args);
    }

    if (POWERBI_TOOL_NAMES.has(name)) {
      if (!(await this.localTools.isAuthenticated())) {
        return signInRequiredResult(name);
      }

      await this.ensureRemoteInitialized();
      return await invokeRemotePowerBITool(this.remote, name, args, this.localTools.defaultContext());
    }

    throw new Error(`Unknown local Power BI tool: ${name}`);
  }

  private async powerBIDiagnostics(): Promise<JSONObject> {
    if (!this.localTools) {
      throw new Error("Local Power BI tools are not configured.");
    }

    const auth = await this.localTools.diagnosticAuthStatus();
    const context = this.localTools.diagnosticSnapshot();
    const localToolNames = this.localTools.tools()
      .map((tool) => tool.name)
      .filter((name): name is string => typeof name === "string");
    const upstream = await this.upstreamDiagnostics(auth.authenticated === true);

    return diagnosticsResult({
      auth,
      context,
      local: {
        toolCount: localToolNames.length,
        expectedCapabilities: toolAvailability(EXPECTED_LOCAL_WRAPPER_TOOLS, new Set(localToolNames)),
      },
      upstream,
    });
  }

  private async upstreamDiagnostics(authenticated: boolean): Promise<JSONObject> {
    if (!authenticated) {
      return {
        checked: false,
        status: "skipped_auth_required",
        advertisedToolCount: 0,
        expectedCapabilities: toolAvailability(EXPECTED_FABRIC_QUERY_TOOLS, new Set()),
      };
    }

    try {
      await this.ensureRemoteInitialized();
      const id = `powerbi-diagnostics-tools-list-${Date.now()}-${this.nextServerRequestId++}`;
      const response = await this.remote.forward({
        jsonrpc: "2.0",
        id,
        method: "tools/list",
      });
      this.captureRemoteTools(response, id);

      const tools = extractRemoteTools(response, id);
      const toolNames = new Set(tools.map((tool) => tool.name));
      return {
        checked: true,
        status: "ok",
        advertisedToolCount: tools.length,
        expectedCapabilities: toolAvailability(EXPECTED_FABRIC_QUERY_TOOLS, toolNames),
      };
    } catch (error) {
      log(`Power BI diagnostics could not list upstream Fabric MCP tools: ${redactedRemoteError(error)}`);
      return {
        checked: true,
        status: "failed",
        error: redactedRemoteError(error),
        advertisedToolCount: 0,
        expectedCapabilities: toolAvailability(EXPECTED_FABRIC_QUERY_TOOLS, new Set()),
      };
    }
  }

  private async prepareRemotePayload(payload: JSONRPCPayload): Promise<RemotePayloadPreparation> {
    if (Array.isArray(payload) || payload.method !== "tools/call" || !isObject(payload.params)) {
      return { payload };
    }

    const name = payload.params.name;
    if (typeof name !== "string") {
      return { payload };
    }

    const args = isObject(payload.params.arguments) ? (payload.params.arguments as JSONObject) : {};
    const tool = await this.remoteTool(name);
    if (!tool || !this.localTools) {
      return { payload };
    }

    const defaulted = applyDefaultArguments(tool, args, this.localTools.defaultContext());
    if (defaulted.status === "missing") {
      return {
        response: {
          jsonrpc: "2.0",
          id: normalizeId(payload.id),
          result: defaulted.result,
        },
      };
    }

    return {
      payload: {
        ...payload,
        params: {
          ...payload.params,
          arguments: defaulted.arguments,
        },
      },
    };
  }

  private async remoteTool(name: string): Promise<RemoteTool | undefined> {
    const cached = this.remoteTools.get(name);
    if (cached) {
      return cached;
    }

    const id = `powerbi-tools-list-${Date.now()}-${this.nextServerRequestId++}`;
    const response = await this.remote.forward({
      jsonrpc: "2.0",
      id,
      method: "tools/list",
    });
    this.captureRemoteTools(response, id);
    return this.remoteTools.get(name);
  }

  private captureRemoteTools(response: JSONValue, id: JSONRPCId): void {
    if (response === null) {
      return;
    }

    const tools = extractRemoteTools(response, id);
    for (const tool of tools) {
      this.remoteTools.set(tool.name, tool);
    }
  }

  private async ensureRemoteInitialized(): Promise<void> {
    if (this.remoteInitialized || !this.initializePayload) {
      return;
    }

    const initializeResponse = await this.remote.forward(this.initializePayload);
    if (initializeResponse !== null) {
      // The local client already received a local initialize response. This response
      // only confirms the upstream session and should not be echoed to Claude.
    }

    this.remoteInitialized = true;
    if (this.initializedNotification) {
      await this.remote.forward(this.initializedNotification);
    }
  }

  private writeServerResponse(response: JSONValue): void {
    if (Array.isArray(response)) {
      for (const item of response) {
        if (isServerMessage(item)) {
          this.writePayload(item);
        }
      }
      return;
    }

    if (isServerMessage(response)) {
      this.writePayload(response);
    }
  }

  private writeForwardError(payload: JSONRPCPayload, error: unknown): void {
    const id = firstRequestId(payload);
    if (id === null && !payloadHasRequestId(payload)) {
      return;
    }

    if (error instanceof RemoteMCPError) {
      this.writePayload(
        jsonRpcError(id, error.code, error.message, {
          http_status: error.httpStatus ?? null,
          details: error.data ?? null,
        }),
      );
      return;
    }

    this.writePayload(jsonRpcError(id, INTERNAL_ERROR, errorMessage(error)));
  }

  private captureInitializeCapabilities(payload: JSONRPCPayload): void {
    const initialize = Array.isArray(payload)
      ? payload.find((item) => isObject(item) && item.method === "initialize")
      : payload;

    if (!isObject(initialize) || initialize.method !== "initialize" || !isObject(initialize.params)) {
      return;
    }

    if (isObject(initialize.params.capabilities)) {
      this.clientCapabilities = initialize.params.capabilities;
    }
  }

  private handleLocalInitialize(payload: JSONRPCPayload): boolean {
    if (Array.isArray(payload) || payload.method !== "initialize") {
      return false;
    }

    this.initializePayload = payload;
    this.remoteInitialized = false;
    this.captureInitializeCapabilities(payload);

    const protocolVersion =
      isObject(payload.params) && typeof payload.params.protocolVersion === "string"
        ? payload.params.protocolVersion
        : "2025-11-25";

    this.writePayload({
      jsonrpc: "2.0",
      id: normalizeId(payload.id),
      result: {
        protocolVersion,
        capabilities: {
          tools: {},
        },
        serverInfo: {
          name: "powerbi-mcp-claude",
          title: "Power BI MCP for Claude",
          version: VERSION,
        },
        instructions:
          "Use the Power BI tools for Microsoft Power BI and Fabric analytics. If a tool says sign-in is required, call powerbi_auth_start and show the returned Microsoft URL and code to the user.",
      },
    });

    return true;
  }

  private handleInitializedNotification(payload: JSONRPCPayload): boolean {
    if (Array.isArray(payload) || payload.method !== "notifications/initialized") {
      return false;
    }

    this.initializedNotification = payload;
    return true;
  }
}

export function makeStdoutWriter(output: Writable): (payload: JSONValue) => void {
  return (payload: JSONValue): void => {
    output.write(`${JSON.stringify(payload)}\n`);
  };
}

type DiagnosticsResultInput = {
  auth: JSONObject;
  context: JSONObject;
  local: JSONObject;
  upstream: JSONObject;
};

function diagnosticsResult(input: DiagnosticsResultInput): JSONObject {
  const context = isObject(input.context.context) ? input.context.context : {};
  const localCapabilities = isObject(input.local.expectedCapabilities) ? input.local.expectedCapabilities : {};
  const upstreamCapabilities = isObject(input.upstream.expectedCapabilities) ? input.upstream.expectedCapabilities : {};
  const upstreamError = typeof input.upstream.error === "string" ? ` (${input.upstream.error})` : "";

  return {
    content: [
      {
        type: "text",
        text: [
          "Power BI MCP diagnostics (redacted):",
          `Server version: ${VERSION}`,
          `Auth: ${String(input.auth.status ?? "unknown")} (authenticated: ${yesNo(input.auth.authenticated)}, cached accounts: ${numberText(input.auth.cachedAccountCount)}, token cache present: ${yesNo(input.auth.tokenCacheExists)})`,
          `Endpoint: configured: ${yesNo(input.context.remoteEndpointConfigured)}, Fabric Power BI MCP endpoint: ${yesNo(input.context.remoteEndpointLooksLikeFabricPowerBI)}`,
          `Context: active workspace: ${yesNo(context.activeWorkspaceConfigured)}, active semantic model: ${yesNo(context.activeSemanticModelConfigured)}, chat workspace override: ${yesNo(context.chatWorkspaceOverrideConfigured)}, chat semantic model override: ${yesNo(context.chatSemanticModelOverrideConfigured)}, default workspace: ${yesNo(context.defaultWorkspaceConfigured)}, default semantic model: ${yesNo(context.defaultSemanticModelConfigured)}`,
          `Local wrappers: ${numberText(input.local.toolCount)} tools; ${capabilitySummary(localCapabilities, EXPECTED_LOCAL_WRAPPER_TOOLS)}`,
          `Upstream Fabric tools: ${String(input.upstream.status ?? "unknown")}${upstreamError}; checked: ${yesNo(input.upstream.checked)}, advertised tools: ${numberText(input.upstream.advertisedToolCount)}; ${capabilitySummary(upstreamCapabilities, EXPECTED_FABRIC_QUERY_TOOLS)}`,
          "Privacy: redacted output omits tokens, auth headers, account email, cache path, and raw workspace or semantic model IDs.",
        ].join("\n"),
      },
    ],
    structuredContent: {
      diagnostics: {
        version: VERSION,
        redacted: true,
        auth: input.auth,
        endpoint: {
          configured: input.context.remoteEndpointConfigured === true,
          looksLikeFabricPowerBI: input.context.remoteEndpointLooksLikeFabricPowerBI === true,
        },
        context,
        local: input.local,
        upstream: input.upstream,
      },
    },
  };
}

function toolAvailability(names: readonly string[], available: Set<string>): JSONObject {
  const capabilities: JSONObject = {};
  for (const name of names) {
    capabilities[name] = available.has(name);
  }
  return capabilities;
}

function capabilitySummary(capabilities: Record<string, unknown>, names: readonly string[]): string {
  return names.map((name) => `${name}: ${yesNo(capabilities[name])}`).join(", ");
}

function yesNo(value: unknown): string {
  return value === true ? "yes" : "no";
}

function numberText(value: unknown): string {
  return typeof value === "number" && Number.isFinite(value) ? String(value) : "unknown";
}

function redactedRemoteError(error: unknown): string {
  if (error instanceof RemoteMCPError) {
    return error.httpStatus ? `Remote MCP HTTP ${error.httpStatus}` : "Remote MCP error";
  }

  return "Remote tool discovery failed";
}

function isServerMessage(value: unknown): value is JSONValue {
  return isObject(value) && ("id" in value || typeof value.method === "string");
}

function payloadHasRequestId(payload: JSONRPCPayload): boolean {
  if (Array.isArray(payload)) {
    return payload.some((item) => isObject(item) && "id" in item && typeof item.method === "string");
  }

  return "id" in payload && typeof payload.method === "string";
}

function formatJsonRpcError(error: unknown): string {
  if (isObject(error) && typeof error.message === "string") {
    return error.message;
  }

  return JSON.stringify(error);
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
