import type { AgentMessage, AgentRequest, CompletionResult, ToolDefinition } from "./agent";

export type { AgentMessage, AgentRequest, CompletionResult, ToolDefinition };

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

export type ChatEventType =
  | "text_delta"
  | "tool_call_start"
  | "tool_call_delta"
  | "tool_result"
  | "done"
  | "error";

export interface ChatEvent {
  type: ChatEventType;
  text?: string;
  toolCallId?: string;
  toolName?: string;
  toolArguments?: string;
  toolResult?: string;
  error?: string;
}

export interface ChatRequest {
  messages: ChatMessage[];
  stream?: boolean;
  maxTokens?: number;
  temperature?: number;
  tools?: ToolDefinition[];
}

export interface PingResult {
  ok: boolean;
  message: string;
  latencyMs?: number;
}

export interface ModelListResult {
  ok: boolean;
  models: string[];
  message: string;
}