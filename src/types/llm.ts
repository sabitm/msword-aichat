export type ProviderKind = "openai" | "anthropic";

export interface ProviderConfig {
  kind: ProviderKind;
  baseUrl: string;
  apiKey: string;
  model: string;
  maxTokens: number;
  temperature: number;
}

export interface ChatMessage {
  role: "system" | "user" | "assistant";
  content: string;
}

export type ChatEventType = "text_delta" | "done" | "error";

export interface ChatEvent {
  type: ChatEventType;
  text?: string;
  error?: string;
}

export interface ChatRequest {
  messages: ChatMessage[];
  stream?: boolean;
  maxTokens?: number;
  temperature?: number;
}

export interface PingResult {
  ok: boolean;
  message: string;
  latencyMs?: number;
}