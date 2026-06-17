import type { UiMessage } from "../hooks/useChat";
import type { AgentStep, PendingEdit } from "../types/agent";

const STORAGE_PREFIX = "msword-aichat:conversation:";
const MAX_STORED_MESSAGES = 80;

interface StoredMessage {
  id: string;
  role: "user" | "assistant";
  content: string;
  error?: string;
  steps?: AgentStep[];
  pendingEdit?: PendingEdit;
}

function storageKey(docKey: string): string {
  return `${STORAGE_PREFIX}${docKey}`;
}

function serializePendingEdit(edit: PendingEdit | undefined): PendingEdit | undefined {
  if (!edit || edit.status === "pending") return undefined;
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
    .filter((message) => !message.isStreaming)
    .slice(-MAX_STORED_MESSAGES)
    .map((message) => ({
      id: message.id,
      role: message.role,
      content: message.content,
      error: message.error,
      steps: message.steps,
      pendingEdit: serializePendingEdit(message.pendingEdit),
    }));
}

function fromStored(stored: StoredMessage[]): UiMessage[] {
  return stored.map((message) => ({
    id: message.id,
    role: message.role,
    content: message.content,
    error: message.error,
    steps: message.steps,
    pendingEdit: message.pendingEdit,
  }));
}

export function loadConversation(docKey: string): UiMessage[] {
  try {
    const raw = localStorage.getItem(storageKey(docKey));
    if (!raw) return [];
    const parsed = JSON.parse(raw) as StoredMessage[];
    if (!Array.isArray(parsed)) return [];
    return fromStored(parsed);
  } catch {
    return [];
  }
}

export function saveConversation(docKey: string, messages: UiMessage[]): void {
  try {
    const payload = toStored(messages);
    localStorage.setItem(storageKey(docKey), JSON.stringify(payload));
  } catch {
    // Ignore quota errors.
  }
}

export function clearConversation(docKey: string): void {
  localStorage.removeItem(storageKey(docKey));
}