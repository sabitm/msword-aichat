import { MessageBar, MessageBarType } from "@fluentui/react";
import * as React from "react";
import type { UiMessage } from "../../hooks/useChat.legacy";
import { useDocumentContext } from "../../hooks/useDocumentContext.legacy";
import { useSettingsStore } from "../../hooks/useSettingsStore.legacy";
import type { ContextMode } from "../../types/context";

import { ChatToolbar } from "./ChatToolbar";
import { MessageInput } from "./MessageInput";
import { MessageList } from "./MessageList";
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

  return (
    <div className="chat-panel">
      <ChatToolbar
        contextMode={props.contextMode}
        context={context}
        isLoading={isLoading}
        docKey={props.docKey}
        persistConversations={persistConversations}
        disabled={!isConfigured || props.isStreaming}
        onContextModeChange={props.onContextModeChange}
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
        <MessageList
          messages={props.messages}
          isStreaming={props.isStreaming}
          onApplyEdit={props.onApplyEdit}
          onRejectEdit={props.onRejectEdit}
          onUndoEdit={props.onUndoEdit}
          onRetry={props.onRetry}
        />
      </div>
      <MessageInput disabled={!isConfigured || props.isStreaming} onSend={props.onSend} />
    </div>
  );
}