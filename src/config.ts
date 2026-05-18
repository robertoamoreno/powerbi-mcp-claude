import { homedir } from "node:os";
import { join } from "node:path";

export const DEFAULT_REMOTE_URL = "https://api.fabric.microsoft.com/v1/mcp/powerbi";
export const DEFAULT_CLIENT_ID = "ea0616ba-638b-4df5-95b9-636659ae5121";
export const DEFAULT_TENANT_ID = "organizations";
export const DEFAULT_AUTHORITY_HOST = "https://login.microsoftonline.com";
export const DEFAULT_SCOPES = ["https://analysis.windows.net/powerbi/api/.default"] as const;
export const MSAL_RESERVED_SCOPES = new Set(["openid", "profile", "offline_access"]);

export type ProxyConfig = {
  remoteUrl: string;
  clientId: string;
  tenantId: string;
  authorityHost: string;
  scopes: string[];
  defaultWorkspaceId?: string;
  defaultWorkspaceName?: string;
  defaultSemanticModelId?: string;
  defaultSemanticModelName?: string;
  cacheDir: string;
  tokenCachePath: string;
  timeoutSeconds: number;
  clientName: string;
  clientVersion: string;
  allowInteractiveAuth: boolean;
  deviceCodeTimeoutSeconds: number;
};

export function loadConfig(env: NodeJS.ProcessEnv = process.env): ProxyConfig {
  const cacheDir =
    env.POWERBI_MCP_CACHE_DIR ??
    (env.CLAUDE_PLUGIN_DATA ? join(env.CLAUDE_PLUGIN_DATA, "cache") : join(homedir(), ".powerbi-mcp-claude"));
  const scopes = parseScopes(env.POWERBI_MCP_SCOPES);

  const config: ProxyConfig = {
    remoteUrl: env.POWERBI_MCP_URL ?? DEFAULT_REMOTE_URL,
    clientId: env.POWERBI_MCP_CLIENT_ID ?? DEFAULT_CLIENT_ID,
    tenantId: env.POWERBI_MCP_TENANT_ID ?? DEFAULT_TENANT_ID,
    authorityHost: env.POWERBI_MCP_AUTHORITY_HOST ?? DEFAULT_AUTHORITY_HOST,
    scopes: scopes.length > 0 ? scopes : [...DEFAULT_SCOPES],
    cacheDir,
    tokenCachePath: join(cacheDir, "msal_token_cache.json"),
    timeoutSeconds: parsePositiveNumber(env.POWERBI_MCP_TIMEOUT_SECONDS, 60),
    clientName: env.POWERBI_MCP_CLIENT_NAME ?? "powerbi-mcp-claude",
    clientVersion: env.POWERBI_MCP_CLIENT_VERSION ?? "0.1.9",
    allowInteractiveAuth: parseBool(env.POWERBI_MCP_ALLOW_INTERACTIVE_AUTH, true),
    deviceCodeTimeoutSeconds: parsePositiveNumber(env.POWERBI_MCP_DEVICE_CODE_TIMEOUT_SECONDS, 900),
  };

  assignOptionalString(config, "defaultWorkspaceId", env.POWERBI_MCP_DEFAULT_WORKSPACE_ID);
  assignOptionalString(config, "defaultWorkspaceName", env.POWERBI_MCP_DEFAULT_WORKSPACE_NAME);
  assignOptionalString(config, "defaultSemanticModelId", env.POWERBI_MCP_DEFAULT_SEMANTIC_MODEL_ID);
  assignOptionalString(config, "defaultSemanticModelName", env.POWERBI_MCP_DEFAULT_SEMANTIC_MODEL_NAME);

  return config;
}

export function authority(config: Pick<ProxyConfig, "authorityHost" | "tenantId">): string {
  return `${config.authorityHost.replace(/\/+$/, "")}/${config.tenantId}`;
}

function parseScopes(raw: string | undefined): string[] {
  if (!raw) {
    return [...DEFAULT_SCOPES];
  }

  return raw
    .replaceAll(",", " ")
    .split(/\s+/)
    .map((scope) => scope.trim())
    .filter((scope) => scope.length > 0 && !MSAL_RESERVED_SCOPES.has(scope));
}

function parsePositiveNumber(raw: string | undefined, fallback: number): number {
  if (!raw) {
    return fallback;
  }

  const parsed = Number(raw);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

function parseBool(raw: string | undefined, fallback: boolean): boolean {
  if (raw === undefined) {
    return fallback;
  }

  return !["0", "false", "no", "off"].includes(raw.trim().toLowerCase());
}

function assignOptionalString<T extends object, K extends keyof T & string>(
  target: T,
  key: K,
  raw: string | undefined,
): void {
  const value = raw?.trim();
  if (value) {
    target[key] = value as T[K];
  }
}
