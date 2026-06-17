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