#!/usr/bin/env node

import { randomBytes } from "node:crypto";
import { rmSync } from "node:fs";
import { createServer, type IncomingMessage, type Server, type ServerResponse } from "node:http";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

import { AuthManager } from "./auth.js";
import { authority, loadConfig, type ProxyConfig } from "./config.js";
import { INTERNAL_ERROR, jsonRpcError, PARSE_ERROR, type JSONValue } from "./jsonrpc.js";
import { log } from "./logger.js";
import { LocalPowerBITools } from "./localTools.js";
import { RemoteMCPClient } from "./remote.js";
import { MCPProxyTransport } from "./transport.js";
import { VERSION } from "./version.js";

const DEFAULT_HTTP_PATH = "/mcp";
const DEFAULT_PORT = 3000;
const DEFAULT_SESSION_TTL_SECONDS = 2 * 60 * 60;
const MAX_REQUEST_BYTES = 2 * 1024 * 1024;

export type HostedMCPServerOptions = {
  config?: ProxyConfig;
  path?: string;
  sessionTtlMs?: number;
  now?: () => number;
  transportFactory?: (config: ProxyConfig) => MCPProxyTransport;
};

export type HostedMCPRuntimeOptions = HostedMCPServerOptions & {
  port?: number;
};

type HostedSession = {
  id: string;
  cacheDir: string;
  config: ProxyConfig;
  transport: MCPProxyTransport;
  createdAt: number;
  touchedAt: number;
};

export class HostedMCPSessionManager {
  private readonly sessions = new Map<string, HostedSession>();

  constructor(
    private readonly baseConfig: ProxyConfig,
    private readonly options: {
      sessionTtlMs: number;
      now: () => number;
      transportFactory: (config: ProxyConfig) => MCPProxyTransport;
    },
  ) {}

  get activeSessionCount(): number {
    return this.sessions.size;
  }

  getSession(requestedSessionId: string | undefined): { session: HostedSession; created: boolean } {
    this.cleanupExpired();

    if (requestedSessionId) {
      const existing = this.sessions.get(requestedSessionId);
      if (existing && !this.isExpired(existing)) {
        existing.touchedAt = this.options.now();
        return { session: existing, created: false };
      }

      if (existing) {
        this.deleteSession(existing.id);
      }
    }

    return { session: this.createSession(), created: true };
  }

  cleanupExpired(): void {
    for (const session of this.sessions.values()) {
      if (this.isExpired(session)) {
        this.deleteSession(session.id);
      }
    }
  }

  deleteSession(sessionId: string): void {
    const session = this.sessions.get(sessionId);
    if (!session) {
      return;
    }

    this.sessions.delete(sessionId);
    try {
      rmSync(session.cacheDir, { recursive: true, force: true });
    } catch (error) {
      log(`Could not remove expired hosted MCP session cache: ${errorMessage(error)}`);
    }
  }

  private createSession(): HostedSession {
    let id = randomSessionId();
    while (this.sessions.has(id)) {
      id = randomSessionId();
    }

    const config = sessionConfig(this.baseConfig, id);
    const now = this.options.now();
    const session: HostedSession = {
      id,
      cacheDir: config.cacheDir,
      config,
      transport: this.options.transportFactory(config),
      createdAt: now,
      touchedAt: now,
    };
    this.sessions.set(id, session);
    return session;
  }

  private isExpired(session: HostedSession): boolean {
    return this.options.now() - session.touchedAt > this.options.sessionTtlMs;
  }
}

export function sessionConfig(baseConfig: ProxyConfig, sessionId: string): ProxyConfig {
  const cacheDir = join(baseConfig.cacheDir, "sessions", sessionId);
  return {
    ...baseConfig,
    cacheDir,
    tokenCachePath: join(cacheDir, "msal_token_cache.json"),
  };
}

export function createHostedMCPTransport(config: ProxyConfig): MCPProxyTransport {
  let transport: MCPProxyTransport | undefined;
  const auth = new AuthManager(config, {
    supportsDeviceCodePrompt: () => transport?.supportsDeviceCodePrompt() ?? false,
    promptDeviceCode: async (response) => {
      if (!transport) {
        return "unavailable";
      }
      return await transport.promptDeviceCode(response);
    },
  });
  const remote = new RemoteMCPClient(config, auth);
  const localTools = new LocalPowerBITools(auth, config);
  transport = new MCPProxyTransport(remote, () => undefined, localTools);
  return transport;
}

export function createHttpMCPServer(options: HostedMCPServerOptions = {}): Server {
  const config = options.config ?? loadConfig();
  const routePath = normalizePath(options.path ?? process.env.POWERBI_MCP_HTTP_PATH ?? DEFAULT_HTTP_PATH);
  const sessionTtlMs = options.sessionTtlMs ?? parsePositiveNumber(process.env.POWERBI_MCP_HTTP_SESSION_TTL_SECONDS, DEFAULT_SESSION_TTL_SECONDS) * 1000;
  const manager = new HostedMCPSessionManager(config, {
    sessionTtlMs,
    now: options.now ?? Date.now,
    transportFactory: options.transportFactory ?? createHostedMCPTransport,
  });

  const server = createServer(async (request, response) => {
    try {
      if (request.method === "GET" && request.url === "/healthz") {
        writeJson(response, 200, {
          ok: true,
          version: VERSION,
          activeSessions: manager.activeSessionCount,
        });
        return;
      }

      if (request.url !== routePath) {
        writeJson(response, 404, { error: "Not found." });
        return;
      }

      if (request.method !== "POST") {
        response.setHeader("Allow", "POST");
        writeJson(response, 405, { error: "Method not allowed." });
        return;
      }

      await handleMCPRequest(request, response, manager);
    } catch (error) {
      log(`Hosted MCP request failed: ${errorMessage(error)}`);
      writeJson(response, 500, jsonRpcError(null, INTERNAL_ERROR, "Internal server error."));
    }
  });

  const cleanupTimer = setInterval(() => manager.cleanupExpired(), Math.min(sessionTtlMs, 60_000));
  cleanupTimer.unref();
  server.on("close", () => clearInterval(cleanupTimer));
  return server;
}

export function loadHostedRuntimeOptions(env: NodeJS.ProcessEnv = process.env): HostedMCPRuntimeOptions {
  return {
    config: loadConfig(env),
    path: normalizePath(env.POWERBI_MCP_HTTP_PATH ?? DEFAULT_HTTP_PATH),
    port: parsePositiveNumber(env.PORT ?? env.POWERBI_MCP_HTTP_PORT, DEFAULT_PORT),
    sessionTtlMs: parsePositiveNumber(env.POWERBI_MCP_HTTP_SESSION_TTL_SECONDS, DEFAULT_SESSION_TTL_SECONDS) * 1000,
  };
}

async function handleMCPRequest(
  request: IncomingMessage,
  response: ServerResponse,
  manager: HostedMCPSessionManager,
): Promise<void> {
  const rawBody = await readRequestBody(request);
  let parsed: unknown;
  try {
    parsed = JSON.parse(rawBody);
  } catch (error) {
    writeJson(response, 400, jsonRpcError(null, PARSE_ERROR, `Parse error: ${errorMessage(error)}`));
    return;
  }

  const requestedSessionId = singleHeader(request.headers["mcp-session-id"]);
  const { session } = manager.getSession(requestedSessionId);
  response.setHeader("Mcp-Session-Id", session.id);

  const responses = await session.transport.handleRequest(parsed);
  if (responses.length === 0) {
    response.statusCode = 202;
    response.end();
    return;
  }

  writeJson(response, 200, responses.length === 1 ? (responses[0] ?? null) : responses);
}

async function readRequestBody(request: IncomingMessage): Promise<string> {
  const chunks: Buffer[] = [];
  let totalBytes = 0;

  for await (const chunk of request) {
    const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
    totalBytes += buffer.byteLength;
    if (totalBytes > MAX_REQUEST_BYTES) {
      throw new Error(`Request body exceeds ${MAX_REQUEST_BYTES} bytes.`);
    }
    chunks.push(buffer);
  }

  return Buffer.concat(chunks).toString("utf8");
}

function writeJson(response: ServerResponse, statusCode: number, payload: JSONValue | Record<string, unknown>): void {
  if (response.headersSent) {
    return;
  }

  response.statusCode = statusCode;
  response.setHeader("Content-Type", "application/json");
  response.end(JSON.stringify(payload));
}

function singleHeader(value: string | string[] | undefined): string | undefined {
  if (Array.isArray(value)) {
    return value[0];
  }
  return value;
}

function normalizePath(path: string): string {
  const trimmed = path.trim();
  if (!trimmed || trimmed === "/") {
    return DEFAULT_HTTP_PATH;
  }
  return trimmed.startsWith("/") ? trimmed : `/${trimmed}`;
}

function parsePositiveNumber(raw: string | undefined, fallback: number): number {
  if (!raw) {
    return fallback;
  }

  const parsed = Number(raw);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

function randomSessionId(): string {
  return randomBytes(32).toString("base64url");
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

async function main(): Promise<void> {
  const options = loadHostedRuntimeOptions();
  const server = createHttpMCPServer(options);
  const port = options.port ?? DEFAULT_PORT;
  const routePath = normalizePath(options.path ?? DEFAULT_HTTP_PATH);
  const config = options.config ?? loadConfig();

  server.listen(port, () => {
    log(`Power BI MCP hosted HTTP server ${VERSION} listening on port ${port}.`);
    log(`MCP endpoint: ${routePath}`);
    log(`Remote Power BI MCP URL: ${config.remoteUrl}`);
    log(`MSAL authority: ${authority(config)}`);
    log(`Hosted session cache root: ${join(config.cacheDir, "sessions")}`);
  });
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  main().catch((error: unknown) => {
    log(errorMessage(error));
    process.exitCode = 1;
  });
}
