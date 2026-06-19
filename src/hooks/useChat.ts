import { useCallback, useEffect, useRef, useState } from "react";
import { buildAgentHistoryFromUi } from "../agent/history";
import { runAgent } from "../agent/orchestrator";
import {
  appendCustomInstructions,
  promptOptionsFromPreferences,
} from "../agent/prompt-options";
import { expandSlashCommand } from "../agent/slash-commands";
import { buildAgentSystemPrompt } from "../agent/system-prompt";
import {
  applyPendingEdit,
  rejectPendingEdit,
  undoPendingEdit,
} from "../agent/tools/registry";
import {
  clearConversation,
  loadConversation,
  saveConversation,
} from "../conversation/store";
import { createProvider } from "../llm/factory";
import type { AppPreferences } from "../settings/defaults";
import { useSettingsStore } from "../settings/store";
import type { AgentMessage, AgentStep, PendingEdit } from "../types/agent";
import type { ContextMode } from "../types/context";
import type { ChatMessage } from "../types/llm";
import { trackEvent } from "../telemetry";
import { buildContextPrompt, getDocumentContext } from "../word/context";

export interface UiMessage {
  id: string;
  role: "user" | "assistant";
  content: string;
  isStreaming?: boolean;
  error?: string;
  steps?: AgentStep[];
  pendingEdit?: PendingEdit;
  agentTranscript?: AgentMessage[];
}

function createId(): string {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

function buildChatSystemPrompt(contextBlock: string | null, preferences: AppPreferences): string {
  const base =
    "You are a helpful writing assistant inside Microsoft Word. Be concise and practical.";

  const withContext = contextBlock
    ? `${base} Use the document context below when answering. If context is empty or missing, say so and ask the user to select text or switch context mode.\n\n${contextBlock}`
    : `${base} Answer using only the conversation unless the user provides document text.`;

  return appendCustomInstructions(withContext, promptOptionsFromPreferences(preferences));
}

export function useChat(contextMode: ContextMode, docKey: string) {
  const getConfig = useSettingsStore((s) => s.getConfig);
  const getPreferences = useSettingsStore((s) => s.getPreferences);
  const [messages, setMessages] = useState<UiMessage[]>([]);
  const [isStreaming, setIsStreaming] = useState(false);
  const loadedDocKeyRef = useRef<string | null>(null);

  useEffect(() => {
    if (docKey === "loading") return;

    if (loadedDocKeyRef.current === docKey) return;
    loadedDocKeyRef.current = docKey;

    const preferences = getPreferences();
    if (!preferences.persistConversations || docKey === "browser") {
      setMessages([]);
      return;
    }

    setMessages(loadConversation(docKey));
  }, [docKey, getPreferences]);

  useEffect(() => {
    if (docKey === "loading" || docKey === "browser" || isStreaming) return;
    if (!getPreferences().persistConversations) return;
    saveConversation(docKey, messages);
  }, [docKey, getPreferences, isStreaming, messages]);

  const updateAssistant = useCallback(
    (assistantId: string, patch: Partial<UiMessage>) => {
      setMessages((prev) =>
        prev.map((message) =>
          message.id === assistantId ? { ...message, ...patch } : message,
        ),
      );
    },
    [],
  );

  const sendMessage = useCallback(
    async (content: string, options?: { skipUserMessage?: boolean }) => {
      const trimmed = content.trim();
      if (!trimmed || isStreaming) return;

      const preferences = getPreferences();
      const slash = expandSlashCommand(trimmed);
      const userVisibleText = slash.displayText;
      const promptText = slash.promptText;

      const userMessage: UiMessage = {
        id: createId(),
        role: "user",
        content: userVisibleText,
      };
      const assistantId = createId();
      const assistantMessage: UiMessage = {
        id: assistantId,
        role: "assistant",
        content: "",
        isStreaming: true,
        steps: preferences.interactionMode === "agent" ? [] : undefined,
      };

      setMessages((prev) =>
        options?.skipUserMessage
          ? [...prev, assistantMessage]
          : [...prev, userMessage, assistantMessage],
      );
      setIsStreaming(true);
      trackEvent(
        preferences.interactionMode === "agent" ? "agent_run" : "chat_message_sent",
        {
          mode: preferences.interactionMode,
          slash: slash.command ?? "",
        },
      );

      const documentContext = await getDocumentContext(contextMode);
      const contextBlock = buildContextPrompt(documentContext);
      const promptOptions = promptOptionsFromPreferences(preferences);

      try {
        if (preferences.interactionMode === "agent") {
          const agentMessages: AgentMessage[] = [
            {
              role: "system",
              content: buildAgentSystemPrompt(contextBlock, promptOptions),
            },
            ...buildAgentHistoryFromUi(messages),
            { role: "user", content: promptText },
          ];

          const result = await runAgent({
            config: getConfig(),
            messages: agentMessages,
            autoApplyEdits: preferences.autoApplyEdits,
            onStep: (step) => {
              setMessages((prev) =>
                prev.map((message) => {
                  if (message.id !== assistantId) return message;
                  const steps = [...(message.steps ?? []), step];
                  return {
                    ...message,
                    steps,
                    content:
                      step.type === "text" && step.text ? step.text : message.content,
                  };
                }),
              );
            },
          });

          updateAssistant(assistantId, {
            content: result.text,
            steps: result.steps,
            pendingEdit: result.pendingEdit,
            agentTranscript: result.transcript,
            isStreaming: false,
            error: result.error,
          });
          return;
        }

        const history: ChatMessage[] = [
          {
            role: "system",
            content: buildChatSystemPrompt(contextBlock, preferences),
          },
          ...messages.map((message) => ({
            role: message.role,
            content: message.content,
          })),
          { role: "user", content: promptText },
        ];

        const provider = createProvider(getConfig());
        for await (const event of provider.chat({ messages: history, stream: true })) {
          if (event.type === "text_delta" && event.text) {
            setMessages((prev) =>
              prev.map((message) =>
                message.id === assistantId
                  ? { ...message, content: message.content + event.text }
                  : message,
              ),
            );
          }

          if (event.type === "error") {
            updateAssistant(assistantId, {
              isStreaming: false,
              error: event.error ?? "Unknown error",
            });
            return;
          }

          if (event.type === "done") {
            updateAssistant(assistantId, { isStreaming: false });
          }
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : "Request failed";
        trackEvent("error_occurred", { surface: "chat", message });
        updateAssistant(assistantId, {
          isStreaming: false,
          error: message,
        });
      } finally {
        setIsStreaming(false);
      }
    },
    [contextMode, getConfig, getPreferences, isStreaming, messages, updateAssistant],
  );

  const applyEdit = useCallback(
    async (messageId: string) => {
      const target = messages.find((message) => message.id === messageId);
      if (!target?.pendingEdit || target.pendingEdit.status !== "pending") return;

      try {
        await applyPendingEdit(target.pendingEdit);
        trackEvent("edit_applied", { tool: target.pendingEdit.toolName });
        setMessages((prev) =>
          prev.map((message) =>
            message.id === messageId && message.pendingEdit
              ? {
                  ...message,
                  pendingEdit: { ...message.pendingEdit, status: "applied" },
                  content: `${message.content}\n\nEdit applied successfully.`,
                }
              : message,
          ),
        );
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Failed to apply edit";
        setMessages((prev) =>
          prev.map((message) =>
            message.id === messageId
              ? { ...message, error: errorMessage }
              : message,
          ),
        );
      }
    },
    [messages],
  );

  const rejectEdit = useCallback(
    async (messageId: string) => {
      const target = messages.find((message) => message.id === messageId);
      if (!target?.pendingEdit || target.pendingEdit.status !== "pending") return;

      try {
        await rejectPendingEdit(target.pendingEdit);
      } catch {
        // Best-effort bookmark cleanup.
      }

      setMessages((prev) =>
        prev.map((message) =>
          message.id === messageId && message.pendingEdit
            ? {
                ...message,
                pendingEdit: { ...message.pendingEdit, status: "rejected" },
                content: `${message.content}\n\nEdit rejected.`,
              }
            : message,
        ),
      );
    },
    [messages],
  );

  const undoEdit = useCallback(
    async (messageId: string) => {
      const target = messages.find((message) => message.id === messageId);
      if (
        !target?.pendingEdit ||
        target.pendingEdit.status !== "applied" ||
        !target.pendingEdit.undo
      ) {
        return;
      }

      try {
        await undoPendingEdit(target.pendingEdit);
        setMessages((prev) =>
          prev.map((message) =>
            message.id === messageId && message.pendingEdit
              ? {
                  ...message,
                  pendingEdit: { ...message.pendingEdit, status: "undone" },
                  content: `${message.content}\n\nEdit undone.`,
                }
              : message,
          ),
        );
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : "Failed to undo edit";
        setMessages((prev) =>
          prev.map((message) =>
            message.id === messageId ? { ...message, error: errorMessage } : message,
          ),
        );
      }
    },
    [messages],
  );

  const clearMessages = useCallback(() => {
    if (isStreaming) return;
    setMessages([]);
    if (docKey !== "loading" && docKey !== "browser") {
      clearConversation(docKey);
    }
  }, [docKey, isStreaming]);

  const retryMessage = useCallback(
    async (assistantMessageId: string) => {
      const assistantIndex = messages.findIndex((message) => message.id === assistantMessageId);
      if (assistantIndex <= 0 || isStreaming) return;

      const userMessage = messages[assistantIndex - 1];
      if (userMessage.role !== "user") return;

      setMessages((prev) => prev.filter((message) => message.id !== assistantMessageId));
      await sendMessage(userMessage.content, { skipUserMessage: true });
    },
    [isStreaming, messages, sendMessage],
  );

  return {
    messages,
    isStreaming,
    sendMessage,
    retryMessage,
    applyEdit,
    rejectEdit,
    undoEdit,
    clearMessages,
  };
}