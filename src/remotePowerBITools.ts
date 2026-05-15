import { isObject, type JSONObject, type JSONRPCPayload, type JSONValue } from "./jsonrpc.js";
import type { RemoteForwarder } from "./remote.js";

type RemoteTool = JSONObject & {
  name: string;
};

export async function invokeRemotePowerBITool(
  remote: RemoteForwarder,
  localName: string,
  args: JSONObject,
): Promise<JSONValue> {
  const toolsListId = nextInternalId("tools-list");
  const toolsResponse = await remote.forward({
    jsonrpc: "2.0",
    id: toolsListId,
    method: "tools/list",
  });

  const tools = extractTools(toolsResponse, toolsListId);
  const tool = findRemoteTool(tools, localName);
  if (!tool) {
    return toolText(
      [
        `The Fabric Power BI MCP endpoint did not advertise a matching upstream tool for ${localName}.`,
        "",
        availableToolsText(tools),
      ].join("\n"),
    );
  }

  const callId = nextInternalId("tools-call");
  const callPayload: JSONRPCPayload = {
    jsonrpc: "2.0",
    id: callId,
    method: "tools/call",
    params: {
      name: tool.name,
      arguments: mapArguments(tool, localName, args),
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

function extractTools(response: JSONValue, id: string): RemoteTool[] {
  const result = extractResult(response, id);
  if (!isObject(result) || !Array.isArray(result.tools)) {
    return [];
  }

  return result.tools.filter((tool): tool is RemoteTool => isObject(tool) && typeof tool.name === "string");
}

function extractResult(response: JSONValue, id: string): JSONValue | undefined {
  const payload = findResponse(response, id);
  if (!payload || !("result" in payload)) {
    return undefined;
  }

  return payload.result as JSONValue;
}

function extractError(response: JSONValue, id: string): unknown {
  const payload = findResponse(response, id);
  if (!payload || !("error" in payload)) {
    return undefined;
  }

  return payload.error;
}

function findResponse(response: JSONValue, id: string): Record<string, unknown> | undefined {
  if (Array.isArray(response)) {
    return response.find((item) => isObject(item) && item.id === id) as Record<string, unknown> | undefined;
  }

  if (isObject(response)) {
    return response;
  }

  return undefined;
}

function findRemoteTool(tools: RemoteTool[], localName: string): RemoteTool | undefined {
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
    if (includesAny(text, ["list", "get", "show", "available", "search"])) {
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
    if (includesAny(text, ["list", "get", "show", "available", "search"])) {
      score += 3;
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
