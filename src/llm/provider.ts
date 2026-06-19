import type { AgentRequest, CompletionResult } from "../types/agent";
import type { ChatEvent, ChatRequest, PingResult, ProviderConfig } from "../types/llm";

export interface LLMProvider {
  readonly kind: ProviderConfig["kind"];
  readonly supportsTools: boolean;
  chat(request: ChatRequest): AsyncIterable<ChatEvent>;
  complete(request: AgentRequest): Promise<CompletionResult>;
  ping(): Promise<PingResult>;
}

export function normalizeBaseUrl(baseUrl: string): string {
  return baseUrl.trim().replace(/\/+$/, "");
}

export function formatNetworkError(error: unknown): string {
  var message = error instanceof Error ? error.message : "Request failed";
  if (
    message === "Failed to fetch" ||
    message.indexOf("NetworkError") >= 0 ||
    message.indexOf("network error") >= 0
  ) {
    return (
      "Failed to fetch — the browser blocked a cross-origin request (CORS). " +
      "Your gateway must answer OPTIONS preflight with 2xx and Access-Control-Allow-* headers, " +
      "or run `npm run proxy` and set base URL to http://localhost:8787/v1."
    );
  }
  return message;
}