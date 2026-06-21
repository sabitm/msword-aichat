import { IconButton, MessageBar, MessageBarType } from "@fluentui/react";
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

var SCROLL_THRESHOLD_PX = 72;

function isNearBottom(element: HTMLElement): boolean {
  return element.scrollHeight - element.scrollTop - element.clientHeight <= SCROLL_THRESHOLD_PX;
}

function scrollPanelToBottom(element: HTMLElement | null): void {
  if (!element) {
    return;
  }
  element.scrollTop = element.scrollHeight;
}

export function ChatPanel(props: ChatPanelProps): React.ReactElement {
  var storeState = useSettingsStore();
  var isConfigured = storeState.isConfigured;
  var persistConversations = storeState.preferences.persistConversations;

  var documentContextState = useDocumentContext(props.contextMode);
  var context = documentContextState.context;
  var isLoading = documentContextState.isLoading;
  var pinAndRefresh = documentContextState.pinAndRefresh;

  var panelBodyRef = React.useRef<HTMLDivElement | null>(null);
  var stickToBottomRef = React.useRef(true);

  var _a = React.useState(false);
  var showJumpToBottom = _a[0];
  var setShowJumpToBottom = _a[1];

  function updateScrollFollowState(): void {
    var panelBody = panelBodyRef.current;
    if (!panelBody) {
      return;
    }
    var nearBottom = isNearBottom(panelBody);
    stickToBottomRef.current = nearBottom;
    setShowJumpToBottom(!nearBottom && props.messages.length > 0);
  }

  function handleJumpToBottom(): void {
    stickToBottomRef.current = true;
    scrollPanelToBottom(panelBodyRef.current);
    setShowJumpToBottom(false);
  }

  function handleSend(message: string): void {
    stickToBottomRef.current = true;
    setShowJumpToBottom(false);
    props.onSend(message);
  }

  React.useLayoutEffect(
    function () {
      if (stickToBottomRef.current) {
        scrollPanelToBottom(panelBodyRef.current);
      }
    },
    [props.messages, props.isStreaming],
  );

  React.useEffect(
    function () {
      stickToBottomRef.current = true;
      setShowJumpToBottom(false);
      scrollPanelToBottom(panelBodyRef.current);
    },
    [props.docKey],
  );

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
          pinAndRefresh();
        }}
      />
      <div className="chat-panel-body-wrap">
        <div
          className="panel-body"
          ref={panelBodyRef}
          onScroll={function () {
            updateScrollFollowState();
          }}
        >
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
        {showJumpToBottom ? (
          <IconButton
            className="jump-to-bottom-btn"
            iconProps={{ iconName: "ChevronDown" }}
            title="Scroll to bottom"
            ariaLabel="Scroll to bottom"
            onClick={handleJumpToBottom}
          />
        ) : null}
      </div>
      <MessageInput disabled={!isConfigured || props.isStreaming} onSend={handleSend} />
    </div>
  );
}