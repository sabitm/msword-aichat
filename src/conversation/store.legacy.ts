import type { UiMessage } from "../hooks/useChat.legacy";
import type { AgentMessage, AgentStep, PendingEdit } from "../types/agent";

var STORAGE_PREFIX = "msword-aichat:conversation:";
var MAX_STORED_MESSAGES = 80;

interface StoredMessage {
  id: string;
  role: "user" | "assistant";
  content: string;
  error?: string;
  steps?: AgentStep[];
  pendingEdit?: PendingEdit;
  agentTranscript?: AgentMessage[];
}

function storageKey(docKey: string): string {
  return STORAGE_PREFIX + docKey;
}

function serializePendingEdit(edit: PendingEdit | undefined): PendingEdit | undefined {
  if (!edit || edit.status === "pending") {
    return undefined;
  }
  return {
    id: edit.id,
    toolName: edit.toolName,
    description: edit.description,
    before: edit.before,
    after: edit.after,
    status: edit.status,
  };
}

function toStored(messages: UiMessage[]): StoredMessage[] {
  return messages
    .filter(function (message) {
      return !message.isStreaming;
    })
    .slice(-MAX_STORED_MESSAGES)
    .map(function (message) {
      return {
        id: message.id,
        role: message.role,
        content: message.content,
        error: message.error,
        steps: message.steps,
        pendingEdit: serializePendingEdit(message.pendingEdit),
        agentTranscript: message.agentTranscript,
      };
    });
}

function fromStored(stored: StoredMessage[]): UiMessage[] {
  return stored.map(function (message) {
    return {
      id: message.id,
      role: message.role,
      content: message.content,
      error: message.error,
      steps: message.steps,
      pendingEdit: message.pendingEdit,
      agentTranscript: message.agentTranscript,
    };
  });
}

export function loadConversation(docKey: string): UiMessage[] {
  try {
    var raw = localStorage.getItem(storageKey(docKey));
    if (!raw) {
      return [];
    }
    var parsed = JSON.parse(raw) as StoredMessage[];
    if (!Array.isArray(parsed)) {
      return [];
    }
    return fromStored(parsed);
  } catch (_error) {
    return [];
  }
}

export function saveConversation(docKey: string, messages: UiMessage[]): void {
  try {
    var payload = toStored(messages);
    localStorage.setItem(storageKey(docKey), JSON.stringify(payload));
  } catch (_error) {
    // Ignore quota errors.
  }
}

export function clearConversation(docKey: string): void {
  localStorage.removeItem(storageKey(docKey));
}