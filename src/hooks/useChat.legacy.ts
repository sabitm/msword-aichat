import * as React from "react";
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
} from "../conversation/store.legacy";
import { createProvider } from "../llm/factory";
import type { AppPreferences } from "../settings/defaults";
import { settingsStore } from "../settings/store.legacy";
import type { AgentMessage, AgentStep, PendingEdit } from "../types/agent";
import type { ContextMode } from "../types/context";
import type { ChatMessage } from "../types/llm";
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
  applyEdit: (messageId: string) => void;
  rejectEdit: (messageId: string) => void;
  undoEdit: (messageId: string) => void;
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
    var slash = expandSlashCommand(trimmed);
    var userVisibleText = slash.displayText;
    var promptText = slash.promptText;

    var userMessage: UiMessage = {
      id: createId(),
      role: "user",
      content: userVisibleText,
    };
    var assistantId = createId();
    var assistantMessage: UiMessage = {
      id: assistantId,
      role: "assistant",
      content: "",
      isStreaming: true,
      steps: preferences.interactionMode === "agent" ? [] : undefined,
    };

    setMessages(function (prev) {
      return options && options.skipUserMessage
        ? prev.concat([assistantMessage])
        : prev.concat([userMessage, assistantMessage]);
    });
    setIsStreaming(true);

    getDocumentContext(contextMode)
      .then(function (documentContext) {
        var contextBlock = buildContextPrompt(documentContext);
        var promptOptions = promptOptionsFromPreferences(preferences);

        if (preferences.interactionMode === "agent") {
          var agentMessages: AgentMessage[] = [
            {
              role: "system",
              content: buildAgentSystemPrompt(contextBlock, promptOptions),
            },
          ];
          var historyMessages = buildAgentHistoryFromUi(messages);
          for (var h = 0; h < historyMessages.length; h++) {
            agentMessages.push(historyMessages[h]);
          }
          agentMessages.push({ role: "user", content: promptText });

          return runAgent({
            config: settingsStore.getConfig(),
            messages: agentMessages,
            autoApplyEdits: preferences.autoApplyEdits,
            onStep: function (step) {
              setMessages(function (prev) {
                return prev.map(function (message) {
                  if (message.id !== assistantId) {
                    return message;
                  }
                  var steps = (message.steps || []).concat([step]);
                  return Object.assign({}, message, {
                    steps: steps,
                    content:
                      step.type === "text" && step.text ? step.text : message.content,
                  });
                });
              });
            },
          }).then(function (result) {
            updateAssistant(assistantId, {
              content: result.text,
              steps: result.steps,
              pendingEdit: result.pendingEdit,
              agentTranscript: result.transcript,
              isStreaming: false,
              error: result.error,
            });
          });
        }

        var history: ChatMessage[] = [
          {
            role: "system",
            content: buildChatSystemPrompt(contextBlock, preferences),
          },
        ];

        for (var j = 0; j < messages.length; j++) {
          history.push({
            role: messages[j].role,
            content: messages[j].content,
          });
        }
        history.push({ role: "user", content: promptText });

        var provider = createProvider(settingsStore.getConfig());
        return consumeChatStream(provider, history, assistantId, updateAssistant, setMessages);
      })
      .catch(function (error) {
        var message = error instanceof Error ? error.message : "Request failed";
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

  function applyEdit(messageId: string): void {
    var target: UiMessage | undefined;
    for (var i = 0; i < messages.length; i++) {
      if (messages[i].id === messageId) {
        target = messages[i];
        break;
      }
    }
    if (!target || !target.pendingEdit || target.pendingEdit.status !== "pending") {
      return;
    }

    applyPendingEdit(target.pendingEdit)
      .then(function () {
        setMessages(function (prev) {
          return prev.map(function (message) {
            if (message.id === messageId && message.pendingEdit) {
              return Object.assign({}, message, {
                pendingEdit: Object.assign({}, message.pendingEdit, { status: "applied" }),
                content: message.content + "\n\nEdit applied successfully.",
              });
            }
            return message;
          });
        });
      })
      .catch(function (error) {
        var errorMessage = error instanceof Error ? error.message : "Failed to apply edit";
        setMessages(function (prev) {
          return prev.map(function (message) {
            return message.id === messageId
              ? Object.assign({}, message, { error: errorMessage })
              : message;
          });
        });
      });
  }

  function rejectEdit(messageId: string): void {
    var target: UiMessage | undefined;
    for (var i = 0; i < messages.length; i++) {
      if (messages[i].id === messageId) {
        target = messages[i];
        break;
      }
    }
    if (!target || !target.pendingEdit || target.pendingEdit.status !== "pending") {
      return;
    }

    rejectPendingEdit(target.pendingEdit)
      .catch(function () {
        // Best-effort bookmark cleanup.
      })
      .then(function () {
        setMessages(function (prev) {
          return prev.map(function (message) {
            if (message.id === messageId && message.pendingEdit) {
              return Object.assign({}, message, {
                pendingEdit: Object.assign({}, message.pendingEdit, { status: "rejected" }),
                content: message.content + "\n\nEdit rejected.",
              });
            }
            return message;
          });
        });
      });
  }

  function undoEdit(messageId: string): void {
    var target: UiMessage | undefined;
    for (var i = 0; i < messages.length; i++) {
      if (messages[i].id === messageId) {
        target = messages[i];
        break;
      }
    }
    if (
      !target ||
      !target.pendingEdit ||
      target.pendingEdit.status !== "applied" ||
      !target.pendingEdit.undo
    ) {
      return;
    }

    undoPendingEdit(target.pendingEdit)
      .then(function () {
        setMessages(function (prev) {
          return prev.map(function (message) {
            if (message.id === messageId && message.pendingEdit) {
              return Object.assign({}, message, {
                pendingEdit: Object.assign({}, message.pendingEdit, { status: "undone" }),
                content: message.content + "\n\nEdit undone.",
              });
            }
            return message;
          });
        });
      })
      .catch(function (error) {
        var errorMessage = error instanceof Error ? error.message : "Failed to undo edit";
        setMessages(function (prev) {
          return prev.map(function (message) {
            return message.id === messageId
              ? Object.assign({}, message, { error: errorMessage })
              : message;
          });
        });
      });
  }

  function retryMessage(assistantMessageId: string): void {
    var assistantIndex = -1;
    for (var k = 0; k < messages.length; k++) {
      if (messages[k].id === assistantMessageId) {
        assistantIndex = k;
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
    applyEdit: applyEdit,
    rejectEdit: rejectEdit,
    undoEdit: undoEdit,
    clearMessages: clearMessages,
  };
}