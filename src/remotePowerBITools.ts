import { isObject, type JSONObject, type JSONRPCId, type JSONRPCPayload, type JSONValue } from "./jsonrpc.js";
import type { PowerBIDefaultContext } from "./localTools.js";
import type { RemoteForwarder } from "./remote.js";

export type RemoteTool = JSONObject & {
  name: string;
};

export type DefaultArgumentResult =
  | {
      status: "ok";
      arguments: JSONObject;
      injected: string[];
    }
  | {
      status: "missing";
      result: JSONObject;
      missingSettings: string[];
    };

const WORKSPACE_FIELDS = [
  "workspaceId",
  "workspace_id",
  "groupId",
  "group_id",
  "workspace",
  "workspaceID",
];

const SEMANTIC_MODEL_FIELDS = [
  "semanticModelId",
  "semantic_model_id",
  "datasetId",
  "dataset_id",
  "artifactId",
  "artifact_id",
];

const REMOTE_TOOL_NAME_BY_LOCAL: Record<string, string> = {
  powerbi_get_semantic_model_schema: "GetSemanticModelSchema",
  powerbi_generate_dax_query: "GenerateQuery",
  powerbi_execute_dax_query: "ExecuteQuery",
  powerbi_get_report_metadata: "GetReportMetadata",
};

export async function invokeRemotePowerBITool(
  remote: RemoteForwarder,
  localName: string,
  args: JSONObject,
  defaultContext: PowerBIDefaultContext,
): Promise<JSONValue> {
  const toolsListId = nextInternalId("tools-list");
  const toolsResponse = await remote.forward({
    jsonrpc: "2.0",
    id: toolsListId,
    method: "tools/list",
  });

  const tools = extractRemoteTools(toolsResponse, toolsListId);
  const tool = findRemoteTool(tools, localName);
  if (!tool) {
    return fallbackResult(localName, defaultContext, tools);
  }

  const defaulted = applyDefaultArguments(
    tool,
    mapArguments(tool, localName, args),
    defaultContext,
  );
  if (defaulted.status === "missing") {
    return defaulted.result;
  }

  const callId = nextInternalId("tools-call");
  const callPayload: JSONRPCPayload = {
    jsonrpc: "2.0",
    id: callId,
    method: "tools/call",
    params: {
      name: tool.name,
      arguments: defaulted.arguments,
    },
  };

  const callResponse = await remote.forward(callPayload);
  const error = extractError(callResponse, callId);
  if (error) {
    return toolText(`Fabric Power BI MCP tool ${tool.name} failed: ${formatRemoteError(error)}`);
  }

  return extractResult(callResponse, callId) ?? toolText(`Fabric Power BI MCP tool ${tool.name} completed without a result.`);
}

export function signInRequiredResult(localName: string): JSONObject {
  return toolText(
    [
      `Power BI sign-in is required before I can run ${localName}.`,
      "Run powerbi_auth_start, open the Microsoft URL, enter the returned code, then run this tool again.",
    ].join("\n"),
  );
}

export function extractRemoteTools(response: JSONValue, id: JSONRPCId): RemoteTool[] {
  const result = extractResult(response, id);
  if (!isObject(result) || !Array.isArray(result.tools)) {
    return [];
  }

  return result.tools.filter((tool): tool is RemoteTool => isObject(tool) && typeof tool.name === "string");
}

export function findRemoteToolByName(tools: RemoteTool[], name: string): RemoteTool | undefined {
  return tools.find((tool) => tool.name === name);
}

export function applyDefaultArguments(
  tool: RemoteTool,
  args: JSONObject,
  defaultContext: PowerBIDefaultContext,
): DefaultArgumentResult {
  const nextArgs: JSONObject = { ...args };
  const injected: string[] = [];
  const missingSettings = new Set<string>();
  const properties = inputPropertyNames(tool);
  const required = requiredInputProperties(tool);

  applyDefaultForFields({
    args: nextArgs,
    properties,
    required,
    fields: WORKSPACE_FIELDS,
    value: defaultContext.workspaceId,
    settingName: "default_workspace_id",
    injected,
    missingSettings,
  });
  applyDefaultForFields({
    args: nextArgs,
    properties,
    required,
    fields: SEMANTIC_MODEL_FIELDS,
    value: defaultContext.semanticModelId,
    settingName: "default_semantic_model_id",
    injected,
    missingSettings,
  });

  if (missingSettings.size > 0) {
    return {
      status: "missing",
      result: missingDefaultContextResult([...missingSettings], tool.name),
      missingSettings: [...missingSettings],
    };
  }

  return {
    status: "ok",
    arguments: nextArgs,
    injected,
  };
}

function extractResult(response: JSONValue, id: JSONRPCId): JSONValue | undefined {
  const payload = findResponse(response, id);
  if (!payload || !("result" in payload)) {
    return undefined;
  }

  return payload.result as JSONValue;
}

function extractError(response: JSONValue, id: JSONRPCId): unknown {
  const payload = findResponse(response, id);
  if (!payload || !("error" in payload)) {
    return undefined;
  }

  return payload.error;
}

function findResponse(response: JSONValue, id: JSONRPCId): Record<string, unknown> | undefined {
  if (Array.isArray(response)) {
    return response.find((item) => isObject(item) && item.id === id) as Record<string, unknown> | undefined;
  }

  if (isObject(response)) {
    return response;
  }

  return undefined;
}

function findRemoteTool(tools: RemoteTool[], localName: string): RemoteTool | undefined {
  const targetName = REMOTE_TOOL_NAME_BY_LOCAL[localName];
  if (targetName) {
    const exact = tools.find((tool) => tool.name.toLowerCase() === targetName.toLowerCase());
    if (exact) {
      return exact;
    }
  }

  const scored = tools
    .map((tool) => ({ tool, score: scoreTool(tool, localName) }))
    .filter((candidate) => candidate.score > 0)
    .sort((left, right) => right.score - left.score);

  const best = scored[0];
  if (!best || best.score < minimumScore(localName)) {
    return undefined;
  }

  return best.tool;
}

function scoreTool(tool: RemoteTool, localName: string): number {
  const text = toolTextForMatching(tool);
  const name = tool.name.toLowerCase();

  if (localName === "powerbi_list_workspaces") {
    let score = 0;
    if (name.includes("workspace")) {
      score += 7;
    }
    if (text.includes("workspace")) {
      score += 5;
    }
    if (text.includes("group")) {
      score += 2;
    }
    if (includesAny(text, ["list", "show", "available", "search"])) {
      score += 3;
    }
    if (includesAny(text, ["semantic", "dataset", "model", "report", "measure"])) {
      score -= 4;
    }
    return score;
  }

  if (localName === "powerbi_list_semantic_models") {
    let score = 0;
    if (text.includes("semantic model") || text.includes("semantic models")) {
      score += 8;
    }
    if (text.includes("dataset") || text.includes("datasets")) {
      score += 7;
    }
    if (text.includes("model") || text.includes("models")) {
      score += 3;
    }
    if (includesAny(text, ["list", "show", "available", "search"])) {
      score += 3;
    }
    if (text.includes("schema")) {
      score -= 8;
    }
    if (text.includes("workspace")) {
      score += 1;
    }
    if (includesAny(text, ["report", "dashboard", "measure"])) {
      score -= 2;
    }
    return score;
  }

  return 0;
}

function minimumScore(localName: string): number {
  return localName === "powerbi_list_workspaces" ? 6 : 7;
}

function toolTextForMatching(tool: RemoteTool): string {
  return [
    tool.name,
    typeof tool.title === "string" ? tool.title : "",
    typeof tool.description === "string" ? tool.description : "",
  ]
    .join(" ")
    .toLowerCase();
}

function includesAny(text: string, values: string[]): boolean {
  return values.some((value) => text.includes(value));
}

function mapArguments(tool: RemoteTool, localName: string, args: JSONObject): JSONObject {
  if (localName === "powerbi_get_semantic_model_schema") {
    return mapSemanticArtifactArguments(args);
  }

  if (localName === "powerbi_generate_dax_query") {
    return mapGenerateQueryArguments(args);
  }

  if (localName === "powerbi_execute_dax_query") {
    return mapExecuteQueryArguments(args);
  }

  if (localName === "powerbi_get_report_metadata") {
    return mapReportMetadataArguments(args);
  }

  if (localName !== "powerbi_list_semantic_models") {
    return {};
  }

  const workspaceId = typeof args.workspaceId === "string" ? args.workspaceId : undefined;
  if (!workspaceId) {
    return {};
  }

  const propertyName = firstInputProperty(tool, [
    "workspaceId",
    "workspace_id",
    "groupId",
    "group_id",
    "workspace",
    "workspaceID",
  ]);

  return propertyName ? { [propertyName]: workspaceId } : { workspaceId };
}

function mapSemanticArtifactArguments(args: JSONObject): JSONObject {
  const mapped: JSONObject = {};
  const artifactId = firstStringArgument(args, [
    "artifactId",
    "artifact_id",
    "semanticModelId",
    "semantic_model_id",
    "datasetId",
    "dataset_id",
  ]);
  if (artifactId) {
    mapped.artifactId = artifactId;
  }
  return mapped;
}

function mapGenerateQueryArguments(args: JSONObject): JSONObject {
  const mapped = mapSemanticArtifactArguments(args);
  const userInput = firstStringArgument(args, ["userInput", "question"]);
  if (userInput) {
    mapped.userInput = userInput;
  }
  if (isObject(args.schemaSelection)) {
    mapped.schemaSelection = args.schemaSelection as JSONValue;
  }
  if (Array.isArray(args.chatHistory)) {
    mapped.chatHistory = args.chatHistory;
  }
  if (Array.isArray(args.valueSearchResults)) {
    mapped.valueSearchResults = args.valueSearchResults;
  }
  return mapped;
}

function mapExecuteQueryArguments(args: JSONObject): JSONObject {
  const mapped = mapSemanticArtifactArguments(args);
  const daxQueries = daxQueriesArgument(args);
  if (daxQueries.length > 0) {
    mapped.daxQueries = daxQueries;
  }
  const maxRows = positiveIntegerArgument(args.maxRows);
  if (maxRows !== undefined) {
    mapped.maxRows = maxRows;
  }
  return mapped;
}

function mapReportMetadataArguments(args: JSONObject): JSONObject {
  const mapped: JSONObject = {};
  const reportObjectId = firstStringArgument(args, ["reportObjectId", "reportId"]);
  if (reportObjectId) {
    mapped.reportObjectId = reportObjectId;
  }
  return mapped;
}

function firstStringArgument(args: JSONObject, names: string[]): string | undefined {
  for (const name of names) {
    const value = args[name];
    if (typeof value !== "string") {
      continue;
    }

    const trimmed = value.trim();
    if (trimmed.length > 0) {
      return trimmed;
    }
  }

  return undefined;
}

function daxQueriesArgument(args: JSONObject): string[] {
  if (Array.isArray(args.daxQueries)) {
    return args.daxQueries
      .filter((value): value is string => typeof value === "string" && value.trim().length > 0)
      .map((value) => value.trim())
      .slice(0, 4);
  }

  const singleQuery = firstStringArgument(args, ["daxQuery", "query"]);
  return singleQuery ? [singleQuery] : [];
}

function positiveIntegerArgument(value: JSONValue | undefined): number | undefined {
  if (typeof value === "number" && Number.isInteger(value) && value > 0) {
    return value;
  }
  if (typeof value !== "string") {
    return undefined;
  }

  const parsed = Number(value);
  if (Number.isInteger(parsed) && parsed > 0) {
    return parsed;
  }

  return undefined;
}

function fallbackResult(
  localName: string,
  defaultContext: PowerBIDefaultContext,
  tools: RemoteTool[],
): JSONObject {
  if (localName === "powerbi_list_workspaces") {
    if (defaultContext.workspaceId) {
      const name = defaultContext.workspaceName ?? "Configured default workspace";
      return {
        content: [
          {
            type: "text",
            text: `Configured Power BI workspace:\n- ${name} (id: ${defaultContext.workspaceId})`,
          },
        ],
        structuredContent: {
          workspaces: [
            {
              id: defaultContext.workspaceId,
              name,
              type: "Configured",
            },
          ],
        },
      };
    }

    return toolText(
      [
        "The Fabric Power BI MCP endpoint did not advertise a workspace listing tool.",
        "Configure default_workspace_id in the Claude Desktop extension settings, then reconnect this MCP server.",
        "",
        availableToolsText(tools),
      ].join("\n"),
    );
  }

  if (localName === "powerbi_list_semantic_models") {
    if (defaultContext.semanticModelId) {
      const workspaceName = defaultContext.workspaceName ?? "Configured default workspace";
      const modelName = defaultContext.semanticModelName ?? "Configured default semantic model";
      const model: JSONObject = {
        id: defaultContext.semanticModelId,
        name: modelName,
        workspaceName,
        workspaceId: defaultContext.workspaceId ?? null,
      };
      return {
        content: [
          {
            type: "text",
            text: [
              "Configured Power BI semantic model:",
              `- ${modelName} (id: ${defaultContext.semanticModelId})`,
              defaultContext.workspaceId ? `Workspace: ${workspaceName} (${defaultContext.workspaceId})` : "Workspace: not configured",
            ].join("\n"),
          },
        ],
        structuredContent: {
          semanticModelGroups: [
            {
              workspaceName,
              workspaceId: defaultContext.workspaceId ?? null,
              models: [model],
            },
          ],
        },
      };
    }

    return toolText(
      [
        "The Fabric Power BI MCP endpoint did not advertise a semantic model listing tool.",
        "Configure default_semantic_model_id in the Claude Desktop extension settings, then reconnect this MCP server.",
        "",
        availableToolsText(tools),
      ].join("\n"),
    );
  }

  return toolText(
    [
      `The Fabric Power BI MCP endpoint did not advertise a matching upstream tool for ${localName}.`,
      "",
      availableToolsText(tools),
    ].join("\n"),
  );
}

function applyDefaultForFields(args: {
  args: JSONObject;
  properties: string[];
  required: string[];
  fields: string[];
  value: string | undefined;
  settingName: string;
  injected: string[];
  missingSettings: Set<string>;
}): void {
  const propertyMatches = args.fields.flatMap((field) =>
    args.properties.filter((property) => property.toLowerCase() === field.toLowerCase()),
  );
  const requiredMatches = args.fields.flatMap((field) =>
    args.required.filter((property) => property.toLowerCase() === field.toLowerCase()),
  );
  const allMatches = [...new Set([...propertyMatches, ...requiredMatches])];

  for (const property of allMatches) {
    if (hasUsableArgument(args.args, property)) {
      continue;
    }

    if (args.value) {
      args.args[property] = args.value;
      args.injected.push(property);
      continue;
    }

    if (requiredMatches.includes(property)) {
      args.missingSettings.add(args.settingName);
    }
  }
}

function hasUsableArgument(args: JSONObject, property: string): boolean {
  const value = args[property];
  return value !== undefined && value !== null && !(typeof value === "string" && value.trim().length === 0);
}

function inputPropertyNames(tool: RemoteTool): string[] {
  if (!isObject(tool.inputSchema) || !isObject(tool.inputSchema.properties)) {
    return [];
  }

  return Object.keys(tool.inputSchema.properties);
}

function requiredInputProperties(tool: RemoteTool): string[] {
  if (!isObject(tool.inputSchema) || !Array.isArray(tool.inputSchema.required)) {
    return [];
  }

  return tool.inputSchema.required.filter((property): property is string => typeof property === "string");
}

function missingDefaultContextResult(settings: string[], toolName: string): JSONObject {
  return toolText(
    [
      `Fabric Power BI MCP tool ${toolName} requires an ID that is not configured.`,
      `Set ${settings.join(" and ")} in the Claude Desktop extension settings, then reconnect this MCP server.`,
    ].join("\n"),
  );
}

function firstInputProperty(tool: RemoteTool, names: string[]): string | undefined {
  if (!isObject(tool.inputSchema) || !isObject(tool.inputSchema.properties)) {
    return undefined;
  }

  const properties = Object.keys(tool.inputSchema.properties);
  for (const name of names) {
    const match = properties.find((property) => property.toLowerCase() === name.toLowerCase());
    if (match) {
      return match;
    }
  }

  return undefined;
}

function availableToolsText(tools: RemoteTool[]): string {
  if (tools.length === 0) {
    return "The endpoint returned no tools.";
  }

  return [
    "Available upstream tools:",
    ...tools.slice(0, 30).map((tool) => {
      const description = typeof tool.description === "string" ? ` - ${tool.description}` : "";
      return `- ${tool.name}${description}`;
    }),
  ].join("\n");
}

function formatRemoteError(error: unknown): string {
  if (isObject(error) && typeof error.message === "string") {
    return error.message;
  }

  return JSON.stringify(error);
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

function nextInternalId(label: string): string {
  return `powerbi-${label}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}
