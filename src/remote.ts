import { type ProxyConfig } from "./config.js";
import { type AuthManager } from "./auth.js";
import { asJsonValue, isObject, type JSONRPCPayload, type JSONValue } from "./jsonrpc.js";
import { log } from "./logger.js";

export class RemoteMCPError extends Error {
  constructor(
    message: string,
    readonly code = -32000,
    readonly httpStatus: number | undefined = undefined,
    readonly data: JSONValue | undefined = undefined,
  ) {
    super(message);
    this.name = "RemoteMCPError";
  }
}

export interface RemoteForwarder {
  forward(payload: JSONRPCPayload): Promise<JSONValue>;
}

export class RemoteMCPClient implements RemoteForwarder {
  private sessionId: string | undefined;
  private protocolVersion: string | undefined;

  constructor(
    private readonly config: ProxyConfig,
    private readonly auth: AuthManager,
  ) {}

  async forward(payload: JSONRPCPayload): Promise<JSONValue> {
    this.captureProtocolVersion(payload);
    return await this.postWithReauth(payload);
  }

  private async postWithReauth(payload: JSONRPCPayload): Promise<JSONValue> {
    let token = await this.auth.getAccessToken();
    let response = await this.post(payload, token);

    if (response.status === 401) {
      log("Remote MCP returned 401; refreshing Microsoft token and retrying once.");
      token = await this.auth.getAccessToken({ forceRefresh: true });
      response = await this.post(payload, token);
    }

    this.captureSessionId(response);

    if (response.status === 202 || response.status === 204) {
      return null;
    }

    if (!response.ok) {
      throw await this.httpError(response);
    }

    const text = await response.text();
    if (!text) {
      return null;
    }

    return parseResponsePayload(text, response.headers.get("content-type") ?? "", response.status);
  }

  private async post(payload: JSONRPCPayload, token: string): Promise<Response> {
    const headers: Record<string, string> = {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      Accept: "application/json, text/event-stream",
      "User-Agent": `${this.config.clientName}/${this.config.clientVersion}`,
    };

    if (this.sessionId) {
      headers["Mcp-Session-Id"] = this.sessionId;
    }
    if (this.protocolVersion) {
      headers["MCP-Protocol-Version"] = this.protocolVersion;
    }

    try {
      return await fetch(this.config.remoteUrl, {
        method: "POST",
        headers,
        body: JSON.stringify(payload),
        signal: AbortSignal.timeout(this.config.timeoutSeconds * 1000),
      });
    } catch (error) {
      throw new RemoteMCPError(`Could not reach remote Power BI MCP server: ${errorMessage(error)}`);
    }
  }

  private captureSessionId(response: Response): void {
    const sessionId = response.headers.get("mcp-session-id");
    if (sessionId && sessionId !== this.sessionId) {
      this.sessionId = sessionId;
      log("Captured remote MCP session id.");
    }
  }

  private captureProtocolVersion(payload: JSONRPCPayload): void {
    const initialize = Array.isArray(payload)
      ? payload.find((item) => isObject(item) && item.method === "initialize")
      : payload;

    if (!isObject(initialize) || initialize.method !== "initialize") {
      return;
    }

    const params = initialize.params;
    if (isObject(params) && typeof params.protocolVersion === "string") {
      this.protocolVersion = params.protocolVersion;
    }
  }

  private async httpError(response: Response): Promise<RemoteMCPError> {
    const text = await response.text();
    let data: JSONValue | undefined;
    let message = text.trim() || response.statusText;

    try {
      const parsed = text ? JSON.parse(text) : undefined;
      data = asJsonValue(parsed);
      if (isObject(parsed)) {
        if (isObject(parsed.error) && typeof parsed.error.message === "string") {
          message = parsed.error.message;
        } else if (typeof parsed.error_description === "string") {
          message = parsed.error_description;
        }
      }
    } catch {
      data = undefined;
    }

    const authChallenge = response.headers.get("www-authenticate");
    if (authChallenge) {
      data = isObject(data)
        ? ({ ...data, www_authenticate: authChallenge } as JSONValue)
        : { www_authenticate: authChallenge };
    }

    return new RemoteMCPError(`Remote MCP HTTP ${response.status}: ${message}`, -32000, response.status, data);
  }
}

export function parseResponsePayload(text: string, contentType: string, status: number): JSONValue {
  if (contentType.includes("text/event-stream")) {
    return parseSsePayload(text);
  }

  try {
    return asJsonValue(JSON.parse(text));
  } catch (error) {
    const preview = text.slice(0, 500).replaceAll("\n", "\\n");
    throw new RemoteMCPError(`Remote MCP returned non-JSON response: ${preview}`, -32000, status);
  }
}

export function parseSsePayload(text: string): JSONValue {
  const events: JSONValue[] = [];
  let dataLines: string[] = [];

  const flush = (): void => {
    if (dataLines.length === 0) {
      return;
    }

    const raw = dataLines.join("\n").trim();
    dataLines = [];
    if (!raw || raw === "[DONE]") {
      return;
    }

    try {
      events.push(asJsonValue(JSON.parse(raw)));
    } catch (error) {
      throw new RemoteMCPError(`Could not parse SSE JSON payload: ${raw.slice(0, 500)}`);
    }
  };

  for (const line of text.split(/\r?\n/)) {
    if (line === "") {
      flush();
    } else if (line.startsWith("data:")) {
      dataLines.push(line.slice(5).trimStart());
    }
  }
  flush();

  if (events.length === 0) {
    return null;
  }
  if (events.length === 1) {
    return events[0] ?? null;
  }
  return events;
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
