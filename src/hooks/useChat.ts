import { useCallback, useState } from "react";
import { createProvider } from "../llm/factory";
import { useSettingsStore } from "../settings/store";
import type { ContextMode } from "../types/context";
import type { ChatMessage } from "../types/llm";
import { buildContextPrompt, getDocumentContext } from "../word/context";

export interface UiMessage {
  id: string;
  role: "user" | "assistant";
  content: string;
  isStreaming?: boolean;
  error?: string;
}

function createId(): string {
  return `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;
}

function buildSystemPrompt(contextBlock: string | null): string {
  const base =
    "You are a helpful writing assistant inside Microsoft Word. Be concise and practical.";

  if (!contextBlock) {
    return `${base} Answer using only the conversation unless the user provides document text.`;
  }

  return `${base} Use the document context below when answering. If context is empty or missing, say so and ask the user to select text or switch context mode.\n\n${contextBlock}`;
}

export function useChat(contextMode: ContextMode) {
  const getConfig = useSettingsStore((s) => s.getConfig);
  const [messages, setMessages] = useState<UiMessage[]>([]);
  const [isStreaming, setIsStreaming] = useState(false);

  const sendMessage = useCallback(
    async (content: string) => {
      const trimmed = content.trim();
      if (!trimmed || isStreaming) return;

      const userMessage: UiMessage = {
        id: createId(),
        role: "user",
        content: trimmed,
      };
      const assistantId = createId();
      const assistantMessage: UiMessage = {
        id: assistantId,
        role: "assistant",
        content: "",
        isStreaming: true,
      };

      setMessages((prev) => [...prev, userMessage, assistantMessage]);
      setIsStreaming(true);

      const documentContext = await getDocumentContext(contextMode);
      const contextBlock = buildContextPrompt(documentContext);

      const history: ChatMessage[] = [
        {
          role: "system",
          content: buildSystemPrompt(contextBlock),
        },
        ...messages.map((m) => ({
          role: m.role,
          content: m.content,
        })),
        { role: "user", content: trimmed },
      ];

      try {
        const provider = createProvider(getConfig());
        for await (const event of provider.chat({ messages: history, stream: true })) {
          if (event.type === "text_delta" && event.text) {
            setMessages((prev) =>
              prev.map((m) =>
                m.id === assistantId
                  ? { ...m, content: m.content + event.text }
                  : m,
              ),
            );
          }

          if (event.type === "error") {
            setMessages((prev) =>
              prev.map((m) =>
                m.id === assistantId
                  ? {
                      ...m,
                      isStreaming: false,
                      error: event.error ?? "Unknown error",
                    }
                  : m,
              ),
            );
            return;
          }

          if (event.type === "done") {
            setMessages((prev) =>
              prev.map((m) =>
                m.id === assistantId ? { ...m, isStreaming: false } : m,
              ),
            );
          }
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : "Request failed";
        setMessages((prev) =>
          prev.map((m) =>
            m.id === assistantId
              ? { ...m, isStreaming: false, error: message }
              : m,
          ),
        );
      } finally {
        setIsStreaming(false);
      }
    },
    [contextMode, getConfig, isStreaming, messages],
  );

  const clearMessages = useCallback(() => {
    if (isStreaming) return;
    setMessages([]);
  }, [isStreaming]);

  return {
    messages,
    isStreaming,
    sendMessage,
    clearMessages,
  };
}