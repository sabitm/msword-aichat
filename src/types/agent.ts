export type InteractionMode = "chat" | "agent";

export interface ToolDefinition {
  name: string;
  description: string;
  parameters: Record<string, unknown>;
}

export interface ToolCall {
  id: string;
  name: string;
  arguments: string;
}

export type AgentStepStatus = "running" | "done" | "error" | "pending";

export interface AgentStep {
  id: string;
  type: "tool_call" | "tool_result" | "text";
  toolName?: string;
  input?: string;
  output?: string;
  text?: string;
  status: AgentStepStatus;
  error?: string;
}

export type PendingEditStatus = "pending" | "applied" | "rejected";

export interface PendingEdit {
  id: string;
  toolName: string;
  description: string;
  before: string;
  after: string;
  status: PendingEditStatus;
}

export interface PendingEditAction {
  toolName: "insert_text" | "replace_text";
  text: string;
  location?: "selection" | "end";
}

export interface ToolExecutionResult {
  success: boolean;
  output: Record<string, unknown>;
  pendingEdit?: PendingEdit;
  error?: string;
}

export interface AgentMessage {
  role: "system" | "user" | "assistant" | "tool";
  content: string;
  toolCalls?: ToolCall[];
  toolCallId?: string;
  name?: string;
}

export interface CompletionResult {
  text: string;
  toolCalls: ToolCall[];
  stopReason: "end" | "tool_calls" | "error";
  error?: string;
}

export interface AgentRequest {
  messages: AgentMessage[];
  tools: ToolDefinition[];
  maxTokens?: number;
  temperature?: number;
}

export interface AgentRunResult {
  text: string;
  steps: AgentStep[];
  pendingEdit?: PendingEdit;
  error?: string;
}

export const MAX_AGENT_STEPS = 10;