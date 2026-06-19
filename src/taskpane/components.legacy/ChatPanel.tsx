import { MessageBar, MessageBarType } from "@fluentui/react";
import * as React from "react";
import type { UiMessage } from "../../hooks/useChat.legacy";
import { useDocumentContext } from "../../hooks/useDocumentContext.legacy";
import { useSettingsStore } from "../../hooks/useSettingsStore.legacy";
import type { ContextMode } from "../../types/context";
import { formatDocumentKeyLabel } from "../../word/document-key.legacy";
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

export function ChatPanel(props: ChatPanelProps): React.ReactElement {
  var storeState = useSettingsStore();
  var isConfigured = storeState.isConfigured;
  var persistConversations = storeState.preferences.persistConversations;

  var documentContextState = useDocumentContext(props.contextMode);
  var context = documentContextState.context;
  var isLoading = documentContextState.isLoading;
  var refresh = documentContextState.refresh;

  var _a = React.useState<string | null>(null);
  var contextNotice = _a[0];
  var setContextNotice = _a[1];

  function handleSend(message: string): void {
    setContextNotice(null);
    props.onSend(message);
  }

  function handleQuickAction(prompt: string): void {
    if (props.contextMode !== "none" && context.empty && !context.error) {
      setContextNotice(
        props.contextMode === "selection"
          ? "Select text in the document first, then try again."
          : "Add headings to the document or switch to Selection context.",
      );
      return;
    }
    setContextNotice(null);
    props.onSend(prompt);
  }

  return (
    <div className="chat-panel">
      <ModeBar />
      {persistConversations && props.docKey !== "loading" && props.docKey !== "browser" ? (
        <div className="doc-session-bar">
          <span>Conversation saved for: {formatDocumentKeyLabel(props.docKey)}</span>
        </div>
      ) : null}
      <ContextBar
        mode={props.contextMode}
        context={context}
        isLoading={isLoading}
        disabled={!isConfigured || props.isStreaming}
        onModeChange={props.onContextModeChange}
        onRefresh={function () {
          refresh();
        }}
      />
      <div className="panel-body">
        {!isConfigured ? (
          <MessageBar messageBarType={MessageBarType.warning} styles={{ root: { margin: 12 } }}>
            Configure your AI endpoint in Settings before chatting.
          </MessageBar>
        ) : null}
        {contextNotice ? (
          <MessageBar
            messageBarType={MessageBarType.warning}
            styles={{ root: { margin: "0 12px 12px" } }}
          >
            {contextNotice}
          </MessageBar>
        ) : null}
        <QuickActions
          contextMode={props.contextMode}
          disabled={!isConfigured || props.isStreaming}
          onAction={handleQuickAction}
        />
        <MessageList
          messages={props.messages}
          isStreaming={props.isStreaming}
          onApplyEdit={props.onApplyEdit}
          onRejectEdit={props.onRejectEdit}
          onUndoEdit={props.onUndoEdit}
          onRetry={props.onRetry}
        />
      </div>
      <MessageInput disabled={!isConfigured || props.isStreaming} onSend={handleSend} />
    </div>
  );
}