export type InteractionMode = "chat" | "agent";

export type MutationToolName =
  | "insert_text"
  | "replace_text"
  | "delete_range"
  | "apply_style"
  | "format_range"
  | "insert_table"
  | "update_table";

export type UndoSnapshotKind = MutationToolName;

export interface UndoSnapshot {
  kind: UndoSnapshotKind;
  bookmark: string;
  previousText: string;
  previousStyle?: DocumentStyleName;
  previousFormat?: {
    bold?: boolean;
    italic?: boolean;
    font_size?: number;
  };
  tableIndex?: number;
  previousTableValues?: string[][];
}

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

export type PendingEditStatus = "pending" | "applied" | "rejected" | "undone";

export interface PendingEdit {
  id: string;
  toolName: MutationToolName | string;
  description: string;
  before: string;
  after: string;
  status: PendingEditStatus;
  bookmark?: string;
  undo?: UndoSnapshot;
  payload?: Record<string, unknown>;
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
  /** Assistant + tool messages produced during this run (for Anthropic multi-turn replay). */
  transcript?: AgentMessage[];
}

export const MAX_AGENT_STEPS = 10;

export const DOCUMENT_STYLES = [
  "Normal",
  "Heading1",
  "Heading2",
  "Heading3",
  "Title",
  "Subtitle",
] as const;

export type DocumentStyleName = (typeof DOCUMENT_STYLES)[number];