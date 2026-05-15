export type JSONValue =
  | null
  | boolean
  | number
  | string
  | JSONValue[]
  | { [key: string]: JSONValue };

export type JSONObject = { [key: string]: JSONValue };
export type JSONRPCId = string | number | null;
export type JSONRPCPayload = JSONObject | JSONValue[];

export const PARSE_ERROR = -32700;
export const INVALID_REQUEST = -32600;
export const INTERNAL_ERROR = -32603;

export function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

export function asJsonValue(value: unknown): JSONValue {
  return value as JSONValue;
}

export function isJsonRpcPayload(value: unknown): value is JSONRPCPayload {
  return isObject(value) || Array.isArray(value);
}

export function isRequest(value: unknown): value is JSONObject & { method: string } {
  return isObject(value) && typeof value.method === "string" && !("result" in value) && !("error" in value);
}

export function isResponse(value: unknown): value is JSONObject & { id: JSONRPCId } {
  return isObject(value) && "id" in value && ("result" in value || "error" in value);
}

export function expectsResponse(payload: JSONRPCPayload): boolean {
  if (Array.isArray(payload)) {
    return payload.some((item) => isObject(item) && "id" in item && typeof item.method === "string");
  }

  return "id" in payload && typeof payload.method === "string";
}

export function firstRequestId(payload: JSONRPCPayload): JSONRPCId {
  if (Array.isArray(payload)) {
    for (const item of payload) {
      if (isObject(item) && "id" in item && typeof item.method === "string") {
        return normalizeId(item.id);
      }
    }
    return null;
  }

  return normalizeId(payload.id);
}

export function normalizeId(value: unknown): JSONRPCId {
  if (typeof value === "string" || typeof value === "number" || value === null) {
    return value;
  }
  return null;
}

export function idKey(id: JSONRPCId): string {
  return `${typeof id}:${String(id)}`;
}

export function jsonRpcError(
  id: JSONRPCId,
  code: number,
  message: string,
  data?: JSONValue,
): JSONObject {
  const error: JSONObject = { code, message };
  if (data !== undefined) {
    error.data = data;
  }

  return {
    jsonrpc: "2.0",
    id,
    error,
  };
}
