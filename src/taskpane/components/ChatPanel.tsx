import { MessageBar, MessageBarBody } from "@fluentui/react-components";
import { useState } from "react";
import type { UiMessage } from "../../hooks/useChat";
import { useDocumentContext } from "../../hooks/useDocumentContext";
import type { ContextMode } from "../../types/context";
import { useSettingsStore } from "../../settings/store";
import { formatDocumentKeyLabel } from "../../word/document-key";
import { ContextBar } from "./ContextBar";
import { MessageInput } from "./MessageInput";
import { MessageList } from "./MessageList";
import { ModeBar } from "./ModeBar";
import { QuickActions } from "./QuickActions";

interface ChatPanelProps {
  messages: UiMessage[];
  isStreaming: boolean;
  docKey: string;
  contextMode: ContextMode;
  onContextModeChange: (mode: ContextMode) => void;
  onSend: (message: string) => void;
  onApplyEdit: (messageId: string) => void;
  onRejectEdit: (messageId: string) => void;
  onUndoEdit: (messageId: string) => void;
  onRetry?: (messageId: string) => void;
}

export function ChatPanel({
  messages,
  isStreaming,
  docKey,
  contextMode,
  onContextModeChange,
  onSend,
  onApplyEdit,
  onRejectEdit,
  onUndoEdit,
  onRetry,
}: ChatPanelProps) {
  const isConfigured = useSettingsStore((s) => s.isConfigured);
  const persistConversations = useSettingsStore((s) => s.preferences.persistConversations);
  const { context, isLoading, refresh } = useDocumentContext(contextMode);
  const [contextNotice, setContextNotice] = useState<string | null>(null);

  const handleSend = (message: string) => {
    setContextNotice(null);
    onSend(message);
  };

  const handleQuickAction = (prompt: string) => {
    if (contextMode !== "none" && context.empty && !context.error) {
      setContextNotice(
        contextMode === "selection"
          ? "Select text in the document first, then try again."
          : "Add headings to the document or switch to Selection context.",
      );
      return;
    }
    setContextNotice(null);
    onSend(prompt);
  };

  return (
    <div className="chat-panel">
      <ModeBar />
      {persistConversations && docKey !== "loading" && docKey !== "browser" ? (
        <div className="doc-session-bar">
          <span>Conversation saved for: {formatDocumentKeyLabel(docKey)}</span>
        </div>
      ) : null}
      <ContextBar
        mode={contextMode}
        context={context}
        isLoading={isLoading}
        disabled={!isConfigured || isStreaming}
        onModeChange={onContextModeChange}
        onRefresh={() => void refresh()}
      />
      <div className="panel-body">
        {!isConfigured ? (
          <MessageBar intent="warning" style={{ margin: 12 }}>
            <MessageBarBody>
              Configure your AI endpoint in Settings before chatting.
            </MessageBarBody>
          </MessageBar>
        ) : null}
        {contextNotice ? (
          <MessageBar intent="warning" style={{ margin: "0 12px 12px" }}>
            <MessageBarBody>{contextNotice}</MessageBarBody>
          </MessageBar>
        ) : null}
        <QuickActions
          contextMode={contextMode}
          disabled={!isConfigured || isStreaming}
          onAction={handleQuickAction}
        />
        <MessageList
          messages={messages}
          isStreaming={isStreaming}
          onApplyEdit={onApplyEdit}
          onRejectEdit={onRejectEdit}
          onUndoEdit={onUndoEdit}
          onRetry={onRetry}
        />
      </div>
      <MessageInput
        disabled={!isConfigured || isStreaming}
        onSend={handleSend}
      />
    </div>
  );
}