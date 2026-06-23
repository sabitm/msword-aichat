import { settingsStore } from "../settings/store.legacy";
import type {
  AgentMessage,
  CompletionResult,
  ToolCall,
  ToolExecutionResult,
} from "../types/agent";

export const AGENT_DEBUG_LOG_STORAGE_KEY = "msword-aichat:agent-debug-log";

export type AgentDebugEventType =
  | "run_start"
  | "llm_request"
  | "llm_response"
  | "tool_call"
  | "tool_result"
  | "run_end";

export interface AgentDebugEvent {
  ts: string;
  runId: string;
  type: AgentDebugEventType;
  stepIndex?: number;
  data: Record<string, unknown>;
}

const MAX_LOG_ENTRIES = 400;
const MAX_LOG_FIELD_CHARS = 6000;

function nowIso(): string {
  return new Date().toISOString();
}

function createId(): string {
  return String(Date.now()) + "-" + Math.random().toString(36).slice(2, 9);
}

export function createAgentRunId(): string {
  return createId();
}

export function isAgentDebugLogEnabled(): boolean {
  return Boolean(settingsStore.getPreferences().agentDebugLogEnabled);
}

function getAgentDebugLogEndpoint(): string {
  return settingsStore.getPreferences().agentDebugLogFileEndpoint.trim();
}

function truncateForLog(value: string, maxChars: number): { text: string; truncated: boolean } {
  if (value.length <= maxChars) {
    return { text: value, truncated: false };
  }
  return {
    text: value.slice(0, maxChars) + "\n...[truncated]",
    truncated: true,
  };
}

function summarizeAgentMessages(messages: AgentMessage[]): Record<string, unknown>[] {
  return messages.map(function (message) {
    var content = truncateForLog(message.content || "", MAX_LOG_FIELD_CHARS);
    var summary: Record<string, unknown> = {
      role: message.role,
      contentLength: message.content ? message.content.length : 0,
      content: content.text,
      contentTruncated: content.truncated,
    };
    if (message.name) {
      summary.name = message.name;
    }
    if (message.toolCallId) {
      summary.toolCallId = message.toolCallId;
    }
    if (message.toolCalls && message.toolCalls.length) {
      summary.toolCalls = message.toolCalls.map(function (call) {
        return {
          id: call.id,
          name: call.name,
          arguments: truncateForLog(call.arguments || "", MAX_LOG_FIELD_CHARS).text,
        };
      });
    }
    return summary;
  });
}

function readStoredEvents(): AgentDebugEvent[] {
  try {
    var raw = localStorage.getItem(AGENT_DEBUG_LOG_STORAGE_KEY);
    if (!raw) {
      return [];
    }
    var parsed = JSON.parse(raw) as AgentDebugEvent[];
    return Array.isArray(parsed) ? parsed : [];
  } catch (_error) {
    return [];
  }
}

function writeStoredEvents(events: AgentDebugEvent[]): void {
  try {
    localStorage.setItem(AGENT_DEBUG_LOG_STORAGE_KEY, JSON.stringify(events));
  } catch (_error) {
    // Best-effort only — quota or private mode.
  }
}

function appendStoredEvent(event: AgentDebugEvent): void {
  var events = readStoredEvents();
  events.push(event);
  if (events.length > MAX_LOG_ENTRIES) {
    events = events.slice(events.length - MAX_LOG_ENTRIES);
  }
  writeStoredEvents(events);
}

function postEventToFileEndpoint(event: AgentDebugEvent): void {
  var endpoint = getAgentDebugLogEndpoint();
  if (!endpoint) {
    return;
  }

  try {
    fetch(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(event),
    }).catch(function () {
      // Dev file logging is optional.
    });
  } catch (_error) {
    // fetch may be unavailable in some hosts.
  }
}

function recordEvent(
  runId: string,
  type: AgentDebugEventType,
  data: Record<string, unknown>,
  stepIndex?: number,
): void {
  if (!isAgentDebugLogEnabled()) {
    return;
  }

  var event: AgentDebugEvent = {
    ts: nowIso(),
    runId: runId,
    type: type,
    data: data,
  };
  if (stepIndex !== undefined) {
    event.stepIndex = stepIndex;
  }

  appendStoredEvent(event);
  postEventToFileEndpoint(event);
}

export function agentDebugLogRunStart(
  runId: string,
  data: { messageCount: number; toolDefinitionCount: number },
): void {
  recordEvent(runId, "run_start", data);
}

export function agentDebugLogLlmRequest(
  runId: string,
  stepIndex: number,
  data: {
    messages: AgentMessage[];
    temperature?: number;
    maxTokens?: number;
  },
): void {
  recordEvent(
    runId,
    "llm_request",
    {
      messageCount: data.messages.length,
      temperature: data.temperature,
      maxTokens: data.maxTokens,
      messages: summarizeAgentMessages(data.messages),
    },
    stepIndex,
  );
}

export function agentDebugLogLlmResponse(
  runId: string,
  stepIndex: number,
  completion: CompletionResult,
): void {
  var text = truncateForLog(completion.text || "", MAX_LOG_FIELD_CHARS);
  recordEvent(
    runId,
    "llm_response",
    {
      stopReason: completion.stopReason,
      error: completion.error,
      textLength: completion.text ? completion.text.length : 0,
      text: text.text,
      textTruncated: text.truncated,
      toolCalls: completion.toolCalls.map(function (call: ToolCall) {
        return {
          id: call.id,
          name: call.name,
          arguments: truncateForLog(call.arguments || "", MAX_LOG_FIELD_CHARS).text,
        };
      }),
    },
    stepIndex,
  );
}

export function agentDebugLogToolCall(
  runId: string,
  stepIndex: number,
  toolCall: ToolCall,
): void {
  var args = truncateForLog(toolCall.arguments || "", MAX_LOG_FIELD_CHARS);
  recordEvent(
    runId,
    "tool_call",
    {
      toolCallId: toolCall.id,
      toolName: toolCall.name,
      arguments: args.text,
      argumentsTruncated: args.truncated,
    },
    stepIndex,
  );
}

export function agentDebugLogToolResult(
  runId: string,
  stepIndex: number,
  toolName: string,
  result: ToolExecutionResult,
): void {
  var outputJson = "";
  try {
    outputJson = JSON.stringify(result.output);
  } catch (_error) {
    outputJson = "[unserializable output]";
  }
  var output = truncateForLog(outputJson, MAX_LOG_FIELD_CHARS);
  recordEvent(
    runId,
    "tool_result",
    {
      toolName: toolName,
      success: result.success,
      error: result.error,
      pendingEdit: Boolean(result.pendingEdit),
      output: output.text,
      outputTruncated: output.truncated,
    },
    stepIndex,
  );
}

export function agentDebugLogRunEnd(
  runId: string,
  data: {
    stepCount: number;
    finalTextLength: number;
    error?: string;
    pendingEdit?: boolean;
  },
): void {
  recordEvent(runId, "run_end", data);
}

export function getAgentDebugLogEntryCount(): number {
  return readStoredEvents().length;
}

export function clearAgentDebugLog(): void {
  try {
    localStorage.removeItem(AGENT_DEBUG_LOG_STORAGE_KEY);
  } catch (_error) {
    // ignore
  }
}

function downloadTextFile(filename: string, text: string): void {
  var blob = new Blob([text], { type: "application/x-ndjson;charset=utf-8" });
  var navigatorRef = window.navigator as Navigator & {
    msSaveOrOpenBlob?: (blob: Blob, filename: string) => void;
  };

  if (navigatorRef.msSaveOrOpenBlob) {
    navigatorRef.msSaveOrOpenBlob(blob, filename);
    return;
  }

  var url = URL.createObjectURL(blob);
  var anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.style.display = "none";
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  URL.revokeObjectURL(url);
}

export function exportAgentDebugLog(): { ok: boolean; message: string } {
  var events = readStoredEvents();
  if (!events.length) {
    return { ok: false, message: "No debug log entries to export." };
  }

  var lines: string[] = [];
  for (var i = 0; i < events.length; i += 1) {
    lines.push(JSON.stringify(events[i]));
  }

  var stamp = nowIso().replace(/[:.]/g, "-");
  downloadTextFile("msword-aichat-agent-debug-" + stamp + ".jsonl", lines.join("\n") + "\n");
  return {
    ok: true,
    message: "Exported " + events.length + " debug log entr" + (events.length === 1 ? "y" : "ies") + ".",
  };
}