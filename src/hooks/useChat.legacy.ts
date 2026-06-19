import * as React from "react";
import {
  appendCustomInstructions,
  promptOptionsFromPreferences,
} from "../agent/prompt-options";
import {
  clearConversation,
  loadConversation,
  saveConversation,
} from "../conversation/store.legacy";
import { createProvider } from "../llm/factory";
import type { AppPreferences } from "../settings/defaults";
import { settingsStore } from "../settings/store.legacy";
import type { AgentStep, PendingEdit } from "../types/agent";
import type { ContextMode } from "../types/context";
import type { ChatMessage } from "../types/llm";
import { trackEvent } from "../telemetry/telemetry.legacy";
import { buildContextPrompt, getDocumentContext } from "../word/context";

export interface UiMessage {
  id: string;
  role: "user" | "assistant";
  content: string;
  isStreaming?: boolean;
  error?: string;
  steps?: AgentStep[];
  pendingEdit?: PendingEdit;
}

function createId(): string {
  return String(Date.now()) + "-" + Math.random().toString(36).slice(2, 9);
}

function buildChatSystemPrompt(contextBlock: string | null, preferences: AppPreferences): string {
  var base =
    "You are a helpful writing assistant inside Microsoft Word. Be concise and practical.";
  var withContext = contextBlock
    ? base +
      " Use the document context below when answering. If context is empty or missing, say so and ask the user to select text or switch context mode.\n\n" +
      contextBlock
    : base + " Answer using only the conversation unless the user provides document text.";

  return appendCustomInstructions(withContext, promptOptionsFromPreferences(preferences));
}

export function useChat(
  contextMode: ContextMode,
  docKey: string,
): {
  messages: UiMessage[];
  isStreaming: boolean;
  sendMessage: (content: string) => void;
  retryMessage: (assistantMessageId: string) => void;
  clearMessages: () => void;
} {
  var _a = React.useState<UiMessage[]>([]);
  var messages = _a[0];
  var setMessages = _a[1];

  var _b = React.useState(false);
  var isStreaming = _b[0];
  var setIsStreaming = _b[1];

  var loadedDocKeyRef = React.useRef<string | null>(null);

  React.useEffect(function () {
    if (docKey === "loading") {
      return;
    }
    if (loadedDocKeyRef.current === docKey) {
      return;
    }
    loadedDocKeyRef.current = docKey;

    var preferences = settingsStore.getPreferences();
    if (!preferences.persistConversations || docKey === "browser") {
      setMessages([]);
      return;
    }

    setMessages(loadConversation(docKey));
  }, [docKey]);

  React.useEffect(function () {
    if (docKey === "loading" || docKey === "browser" || isStreaming) {
      return;
    }
    if (!settingsStore.getPreferences().persistConversations) {
      return;
    }
    saveConversation(docKey, messages);
  }, [docKey, isStreaming, messages]);

  function updateAssistant(assistantId: string, patch: Partial<UiMessage>): void {
    setMessages(function (prev) {
      return prev.map(function (message) {
        return message.id === assistantId ? Object.assign({}, message, patch) : message;
      });
    });
  }

  function sendMessage(content: string, options?: { skipUserMessage?: boolean }): void {
    var trimmed = content.trim();
    if (!trimmed || isStreaming) {
      return;
    }

    var preferences = settingsStore.getPreferences();
    if (preferences.interactionMode === "agent") {
      var agentAssistantId = createId();
      var agentNotice: UiMessage = {
        id: agentAssistantId,
        role: "assistant",
        content:
          "Agent mode arrives in IE-3. Open Settings and switch interaction mode to Chat, or wait for the next update.",
        error: undefined,
      };
      var agentUser: UiMessage = {
        id: createId(),
        role: "user",
        content: trimmed,
      };
      setMessages(function (prev) {
        return options && options.skipUserMessage
          ? prev.concat([agentNotice])
          : prev.concat([agentUser, agentNotice]);
      });
      return;
    }

    var userMessage: UiMessage = {
      id: createId(),
      role: "user",
      content: trimmed,
    };
    var assistantId = createId();
    var assistantMessage: UiMessage = {
      id: assistantId,
      role: "assistant",
      content: "",
      isStreaming: true,
    };

    setMessages(function (prev) {
      return options && options.skipUserMessage
        ? prev.concat([assistantMessage])
        : prev.concat([userMessage, assistantMessage]);
    });
    setIsStreaming(true);
    trackEvent("chat_message_sent", { mode: "chat" });

    getDocumentContext(contextMode)
      .then(function (documentContext) {
        var contextBlock = buildContextPrompt(documentContext);
        var history: ChatMessage[] = [
          {
            role: "system",
            content: buildChatSystemPrompt(contextBlock, preferences),
          },
        ];

        for (var i = 0; i < messages.length; i++) {
          history.push({
            role: messages[i].role,
            content: messages[i].content,
          });
        }
        history.push({ role: "user", content: trimmed });

        var provider = createProvider(settingsStore.getConfig());
        return consumeChatStream(provider, history, assistantId, updateAssistant, setMessages);
      })
      .catch(function (error) {
        var message = error instanceof Error ? error.message : "Request failed";
        trackEvent("error_occurred", { surface: "chat", message: message });
        updateAssistant(assistantId, {
          isStreaming: false,
          error: message,
        });
      })
      .then(function () {
        setIsStreaming(false);
      });
  }

  async function consumeChatStream(
    provider: ReturnType<typeof createProvider>,
    history: ChatMessage[],
    assistantId: string,
    updateFn: (assistantId: string, patch: Partial<UiMessage>) => void,
    setMessagesFn: React.Dispatch<React.SetStateAction<UiMessage[]>>,
  ): Promise<void> {
    for await (const event of provider.chat({ messages: history, stream: true })) {
      if (event.type === "text_delta" && event.text) {
        setMessagesFn(function (prev) {
          return prev.map(function (message) {
            return message.id === assistantId
              ? Object.assign({}, message, { content: message.content + event.text })
              : message;
          });
        });
      }

      if (event.type === "error") {
        updateFn(assistantId, {
          isStreaming: false,
          error: event.error || "Unknown error",
        });
        return;
      }

      if (event.type === "done") {
        updateFn(assistantId, { isStreaming: false });
        return;
      }
    }

    updateFn(assistantId, { isStreaming: false });
  }

  function retryMessage(assistantMessageId: string): void {
    var assistantIndex = -1;
    for (var i = 0; i < messages.length; i++) {
      if (messages[i].id === assistantMessageId) {
        assistantIndex = i;
        break;
      }
    }
    if (assistantIndex <= 0 || isStreaming) {
      return;
    }

    var userMessage = messages[assistantIndex - 1];
    if (userMessage.role !== "user") {
      return;
    }

    setMessages(function (prev) {
      return prev.filter(function (message) {
        return message.id !== assistantMessageId;
      });
    });
    sendMessage(userMessage.content, { skipUserMessage: true });
  }

  function clearMessages(): void {
    if (isStreaming) {
      return;
    }
    setMessages([]);
    if (docKey !== "loading" && docKey !== "browser") {
      clearConversation(docKey);
    }
  }

  return {
    messages: messages,
    isStreaming: isStreaming,
    sendMessage: sendMessage,
    retryMessage: retryMessage,
    clearMessages: clearMessages,
  };
}