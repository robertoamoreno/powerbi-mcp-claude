import type { AuthManager } from "./auth.js";
import { isObject, type JSONObject, type JSONValue } from "./jsonrpc.js";

export const AUTH_TOOL_NAMES = new Set([
  "powerbi_auth_start",
  "powerbi_auth_status",
  "powerbi_auth_logout",
]);

export const POWERBI_TOOL_NAMES = new Set([
  "powerbi_list_workspaces",
  "powerbi_list_semantic_models",
]);

const LOCAL_TOOL_NAMES = new Set([...AUTH_TOOL_NAMES, ...POWERBI_TOOL_NAMES]);

export class LocalPowerBITools {
  constructor(private readonly auth: AuthManager) {}

  async isAuthenticated(): Promise<boolean> {
    return (await this.auth.getCachedAccessToken()) !== undefined;
  }

  tools(): JSONObject[] {
    return [
      {
        name: "powerbi_auth_start",
        title: "Start Power BI sign-in",
        description:
          "Start Microsoft device-code sign-in for Power BI. Returns the Microsoft URL and code to show the user in chat.",
        inputSchema: {
          type: "object",
          properties: {},
          additionalProperties: false,
        },
      },
      {
        name: "powerbi_auth_status",
        title: "Check Power BI sign-in",
        description: "Check whether Microsoft Power BI authentication has completed.",
        inputSchema: {
          type: "object",
          properties: {},
          additionalProperties: false,
        },
      },
      {
        name: "powerbi_auth_logout",
        title: "Sign out of Power BI",
        description: "Delete the local Microsoft token cache used by this Power BI MCP extension.",
        inputSchema: {
          type: "object",
          properties: {},
          additionalProperties: false,
        },
      },
      {
        name: "powerbi_list_workspaces",
        title: "List Power BI workspaces",
        description:
          "List the Power BI workspaces available to the signed-in Microsoft account.",
        inputSchema: {
          type: "object",
          properties: {},
          additionalProperties: false,
        },
      },
      {
        name: "powerbi_list_semantic_models",
        title: "List Power BI semantic models",
        description:
          "List Power BI semantic models, also known as datasets. If no workspaceId is provided, lists My workspace and every accessible workspace.",
        inputSchema: {
          type: "object",
          properties: {
            workspaceId: {
              type: "string",
              description:
                "Optional Power BI workspace ID. Omit it to list semantic models from My workspace and all accessible workspaces.",
            },
          },
          additionalProperties: false,
        },
      },
    ];
  }

  async handleToolCall(name: string, args: JSONObject = {}): Promise<JSONValue> {
    if (name === "powerbi_auth_start") {
      const result = await this.auth.startDeviceCodeLogin();
      if (result.status === "authenticated") {
        return toolText(result.message);
      }

      return toolText(
        [
          "Microsoft sign-in is required for Power BI.",
          "",
          `Open: ${result.verificationUri}`,
          `Code: ${result.userCode}`,
          "",
          "After sign-in completes in the browser, ask me to check Power BI sign-in status.",
        ].join("\n"),
      );
    }

    if (name === "powerbi_auth_status") {
      const status = await this.auth.deviceLoginStatus();
      const info = await this.auth.tokenInfo();
      return toolText(
        [
          status.message,
          `Token cache: ${info.cachePath}`,
          `Cached accounts: ${info.accountCount}`,
        ].join("\n"),
      );
    }

    if (name === "powerbi_auth_logout") {
      const result = await this.auth.clearCache();
      return toolText(
        [
          "Signed out of Power BI locally and cleared the Microsoft token cache.",
          `Token cache: ${result.cachePath}`,
          `Accounts removed: ${result.accountsRemoved}`,
          `Remaining cached accounts: ${result.remainingAccounts}`,
        ].join("\n"),
      );
    }

    throw new Error(`Unknown local Power BI tool: ${name}`);
  }
}

export function isLocalToolCall(params: unknown): string | undefined {
  if (!isObject(params)) {
    return undefined;
  }

  const name = params.name;
  if (typeof name === "string" && LOCAL_TOOL_NAMES.has(name)) {
    return name;
  }

  return undefined;
}

export function mergeTools(remoteResponse: JSONValue, localTools: JSONObject[]): JSONValue {
  if (!isObject(remoteResponse) || !isObject(remoteResponse.result)) {
    return remoteResponse;
  }

  const remoteTools = Array.isArray(remoteResponse.result.tools) ? remoteResponse.result.tools : [];
  const localToolNames = new Set(localTools.map((tool) => tool.name).filter((name): name is string => typeof name === "string"));
  const uniqueRemoteTools = remoteTools.filter((tool) => {
    if (!isObject(tool) || typeof tool.name !== "string") {
      return true;
    }
    return !localToolNames.has(tool.name);
  });

  return {
    ...remoteResponse,
    result: {
      ...remoteResponse.result,
      tools: [...localTools, ...uniqueRemoteTools],
    },
  };
}

export function toolsListResult(id: JSONValue, tools: JSONObject[]): JSONObject {
  return {
    jsonrpc: "2.0",
    id,
    result: {
      tools,
    },
  };
}

function toolText(text: string): JSONObject {
  return {
    content: [
      {
        type: "text",
        text,
      },
    ],
  };
}
