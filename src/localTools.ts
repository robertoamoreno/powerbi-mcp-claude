import type { AuthManager } from "./auth.js";
import type { ProxyConfig } from "./config.js";
import { isObject, type JSONObject, type JSONValue } from "./jsonrpc.js";

export const AUTH_TOOL_NAMES = new Set([
  "powerbi_auth_start",
  "powerbi_auth_status",
  "powerbi_auth_logout",
]);

export const CONFIG_TOOL_NAMES = new Set([
  "powerbi_get_default_context",
  "powerbi_set_context",
  "powerbi_clear_context",
]);

export const POWERBI_TOOL_NAMES = new Set([
  "powerbi_list_workspaces",
  "powerbi_list_semantic_models",
  "powerbi_get_semantic_model_schema",
  "powerbi_generate_dax_query",
  "powerbi_execute_dax_query",
  "powerbi_get_report_metadata",
]);

const LOCAL_TOOL_NAMES = new Set([...AUTH_TOOL_NAMES, ...CONFIG_TOOL_NAMES, ...POWERBI_TOOL_NAMES]);

export type PowerBIDefaultContext = {
  workspaceId?: string;
  workspaceName?: string;
  semanticModelId?: string;
  semanticModelName?: string;
};

export class LocalPowerBITools {
  private chatContext: PowerBIDefaultContext = {};

  constructor(
    private readonly auth: AuthManager,
    private readonly config: Pick<
      ProxyConfig,
      "defaultWorkspaceId" | "defaultWorkspaceName" | "defaultSemanticModelId" | "defaultSemanticModelName"
    > = {},
  ) {}

  async isAuthenticated(): Promise<boolean> {
    return (await this.auth.getCachedAccessToken()) !== undefined;
  }

  defaultContext(): PowerBIDefaultContext {
    return mergeContexts(this.configuredContext(), this.chatContext);
  }

  private configuredContext(): PowerBIDefaultContext {
    const context: PowerBIDefaultContext = {};
    if (this.config.defaultWorkspaceId) {
      context.workspaceId = this.config.defaultWorkspaceId;
    }
    if (this.config.defaultWorkspaceName) {
      context.workspaceName = this.config.defaultWorkspaceName;
    }
    if (this.config.defaultSemanticModelId) {
      context.semanticModelId = this.config.defaultSemanticModelId;
    }
    if (this.config.defaultSemanticModelName) {
      context.semanticModelName = this.config.defaultSemanticModelName;
    }
    return context;
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
        name: "powerbi_get_default_context",
        title: "Get Power BI default context",
        description:
          "Show the configured default Power BI workspace and semantic model IDs used when Fabric MCP tools require IDs.",
        inputSchema: {
          type: "object",
          properties: {},
          additionalProperties: false,
        },
      },
      {
        name: "powerbi_set_context",
        title: "Set Power BI context",
        description:
          "Set the active Power BI workspace and semantic model IDs for this chat session. These values are not persisted and override Claude Desktop defaults until cleared.",
        inputSchema: {
          type: "object",
          properties: {
            workspaceId: {
              type: "string",
              description: "Power BI workspace ID to use for later tool calls in this chat session.",
            },
            workspaceName: {
              type: "string",
              description: "Optional display name for the active workspace.",
            },
            semanticModelId: {
              type: "string",
              description: "Power BI semantic model, dataset, or artifact ID to use for later tool calls in this chat session.",
            },
            semanticModelName: {
              type: "string",
              description: "Optional display name for the active semantic model.",
            },
          },
          additionalProperties: false,
        },
      },
      {
        name: "powerbi_clear_context",
        title: "Clear Power BI context",
        description:
          "Clear the active chat-session Power BI context and go back to the Claude Desktop default settings.",
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
      {
        name: "powerbi_get_semantic_model_schema",
        title: "Get semantic model schema",
        description:
          "Get Power BI semantic model metadata from Fabric MCP. Uses the active semantic model context when no semanticModelId or artifactId is supplied.",
        inputSchema: {
          type: "object",
          properties: {
            semanticModelId: {
              type: "string",
              description: "Optional semantic model ID. Omit to use the active Power BI context.",
            },
            artifactId: {
              type: "string",
              description: "Optional Fabric artifact ID. Omit to use the active Power BI context.",
            },
          },
          additionalProperties: false,
        },
      },
      {
        name: "powerbi_generate_dax_query",
        title: "Generate DAX query",
        description:
          "Ask Fabric MCP to generate a DAX query for the active semantic model. Uses the active semantic model context when no semanticModelId or artifactId is supplied.",
        inputSchema: {
          type: "object",
          properties: {
            userInput: {
              type: "string",
              description: "The business question to generate DAX for.",
            },
            question: {
              type: "string",
              description: "Alias for userInput.",
            },
            semanticModelId: {
              type: "string",
              description: "Optional semantic model ID. Omit to use the active Power BI context.",
            },
            artifactId: {
              type: "string",
              description: "Optional Fabric artifact ID. Omit to use the active Power BI context.",
            },
            schemaSelection: {
              type: "object",
              description: "Optional subset of tables, columns, and measures relevant to the question.",
            },
            chatHistory: {
              type: "array",
              description: "Optional prior user/assistant messages for DAX generation context.",
            },
            valueSearchResults: {
              type: "array",
              description: "Optional resolved value matches for accurate filters.",
            },
          },
          additionalProperties: false,
        },
      },
      {
        name: "powerbi_execute_dax_query",
        title: "Execute DAX query",
        description:
          "Execute one or more DAX queries against the active Power BI semantic model through Fabric MCP. Uses the active semantic model context when no semanticModelId or artifactId is supplied.",
        inputSchema: {
          type: "object",
          properties: {
            daxQuery: {
              type: "string",
              description: "A single DAX query with one EVALUATE statement.",
            },
            daxQueries: {
              type: "array",
              items: {
                type: "string",
              },
              description: "One to four DAX queries, each with one EVALUATE statement.",
            },
            semanticModelId: {
              type: "string",
              description: "Optional semantic model ID. Omit to use the active Power BI context.",
            },
            artifactId: {
              type: "string",
              description: "Optional Fabric artifact ID. Omit to use the active Power BI context.",
            },
            maxRows: {
              type: "integer",
              description: "Optional maximum rows per query. Fabric MCP defaults to 250 and supports up to 1000.",
              minimum: 1,
              maximum: 1000,
            },
          },
          additionalProperties: false,
        },
      },
      {
        name: "powerbi_get_report_metadata",
        title: "Get Power BI report metadata",
        description:
          "Get high-level Power BI report metadata from Fabric MCP, including report pages, visuals, filters, and semantic model details.",
        inputSchema: {
          type: "object",
          properties: {
            reportObjectId: {
              type: "string",
              description: "Power BI report object ID to inspect.",
            },
          },
          required: ["reportObjectId"],
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

    if (name === "powerbi_get_default_context") {
      return defaultContextResult({
        heading: "Power BI active context:",
        activeContext: this.defaultContext(),
        chatContext: this.chatContext,
        configuredContext: this.configuredContext(),
      });
    }

    if (name === "powerbi_set_context") {
      const nextContext = contextFromArgs(args);
      if (!hasContextValue(nextContext)) {
        return toolText(
          "Provide at least one context value: workspaceId, workspaceName, semanticModelId, or semanticModelName.",
        );
      }

      this.chatContext = mergeContexts(this.chatContext, nextContext);
      return defaultContextResult({
        heading: "Updated Power BI chat context:",
        activeContext: this.defaultContext(),
        chatContext: this.chatContext,
        configuredContext: this.configuredContext(),
      });
    }

    if (name === "powerbi_clear_context") {
      this.chatContext = {};
      return defaultContextResult({
        heading: "Cleared Power BI chat context:",
        activeContext: this.defaultContext(),
        chatContext: this.chatContext,
        configuredContext: this.configuredContext(),
      });
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

type DefaultContextResultInput = {
  heading: string;
  activeContext: PowerBIDefaultContext;
  chatContext: PowerBIDefaultContext;
  configuredContext: PowerBIDefaultContext;
};

function defaultContextResult(input: DefaultContextResultInput): JSONObject {
  return {
    content: [
      {
        type: "text",
        text: [
          input.heading,
          `Active workspace: ${contextLabel(input.activeContext, "workspace")}`,
          `Active semantic model: ${contextLabel(input.activeContext, "semanticModel")}`,
          "",
          `Chat workspace override: ${contextLabel(input.chatContext, "workspace")}`,
          `Chat semantic model override: ${contextLabel(input.chatContext, "semanticModel")}`,
          "",
          `Configured workspace default: ${contextLabel(input.configuredContext, "workspace")}`,
          `Configured semantic model default: ${contextLabel(input.configuredContext, "semanticModel")}`,
        ].join("\n"),
      },
    ],
    structuredContent: {
      defaultContext: structuredContext(input.activeContext),
      activeContext: structuredContext(input.activeContext),
      chatContext: structuredContext(input.chatContext),
      configuredContext: structuredContext(input.configuredContext),
    },
  };
}

function contextFromArgs(args: JSONObject): PowerBIDefaultContext {
  const context: PowerBIDefaultContext = {};
  assignStringArg(context, "workspaceId", args.workspaceId);
  assignStringArg(context, "workspaceName", args.workspaceName);
  assignStringArg(context, "semanticModelId", args.semanticModelId);
  assignStringArg(context, "semanticModelName", args.semanticModelName);
  return context;
}

function assignStringArg(
  context: PowerBIDefaultContext,
  key: keyof PowerBIDefaultContext,
  value: JSONValue | undefined,
): void {
  if (typeof value !== "string") {
    return;
  }

  const trimmed = value.trim();
  if (trimmed.length === 0) {
    return;
  }

  context[key] = trimmed;
}

function mergeContexts(base: PowerBIDefaultContext, override: PowerBIDefaultContext): PowerBIDefaultContext {
  const context: PowerBIDefaultContext = {};
  assignContextValue(context, "workspaceId", base.workspaceId);
  assignContextValue(context, "workspaceName", base.workspaceName);
  assignContextValue(context, "semanticModelId", base.semanticModelId);
  assignContextValue(context, "semanticModelName", base.semanticModelName);
  assignContextValue(context, "workspaceId", override.workspaceId);
  assignContextValue(context, "workspaceName", override.workspaceName);
  assignContextValue(context, "semanticModelId", override.semanticModelId);
  assignContextValue(context, "semanticModelName", override.semanticModelName);
  return context;
}

function assignContextValue(
  context: PowerBIDefaultContext,
  key: keyof PowerBIDefaultContext,
  value: string | undefined,
): void {
  if (value) {
    context[key] = value;
  }
}

function hasContextValue(context: PowerBIDefaultContext): boolean {
  return Boolean(
    context.workspaceId ||
      context.workspaceName ||
      context.semanticModelId ||
      context.semanticModelName,
  );
}

function contextLabel(context: PowerBIDefaultContext, kind: "workspace" | "semanticModel"): string {
  if (kind === "workspace") {
    if (context.workspaceId) {
      return `${context.workspaceName ?? "Workspace"} (${context.workspaceId})`;
    }
    return context.workspaceName ?? "Not configured";
  }

  if (context.semanticModelId) {
    return `${context.semanticModelName ?? "Semantic model"} (${context.semanticModelId})`;
  }
  return context.semanticModelName ?? "Not configured";
}

function structuredContext(context: PowerBIDefaultContext): JSONObject {
  return {
    workspaceId: context.workspaceId ?? null,
    workspaceName: context.workspaceName ?? null,
    semanticModelId: context.semanticModelId ?? null,
    semanticModelName: context.semanticModelName ?? null,
  };
}
