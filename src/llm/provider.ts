import type { ChatEvent, ChatRequest, PingResult, ProviderConfig } from "../types/llm";

export interface LLMProvider {
  readonly kind: ProviderConfig["kind"];
  chat(request: ChatRequest): AsyncIterable<ChatEvent>;
  ping(): Promise<PingResult>;
}

export function normalizeBaseUrl(baseUrl: string): string {
  return baseUrl.trim().replace(/\/+$/, "");
}