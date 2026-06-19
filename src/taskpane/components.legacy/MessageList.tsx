import { Spinner, SpinnerSize, Text } from "@fluentui/react";
import * as React from "react";
import type { UiMessage } from "../../hooks/useChat.legacy";
import { AgentTrace } from "./AgentTrace";
import { EditPreview } from "./EditPreview";
import { ErrorActions } from "./ErrorActions";

interface MessageListProps {
  messages: UiMessage[];
  isStreaming: boolean;
  onApplyEdit: (messageId: string) => void;
  onRejectEdit: (messageId: string) => void;
  onUndoEdit: (messageId: string) => void;
  onRetry?: (messageId: string) => void;
}

export function MessageList(props: MessageListProps): React.ReactElement {
  if (props.messages.length === 0) {
    return (
      <div className="empty-state">
        <Text variant="large" block>
          Start a conversation
        </Text>
        <Text variant="medium" block>
          Ask questions about your document, or switch to Agent mode to edit with tools.
        </Text>
      </div>
    );
  }

  return (
    <div className="message-list">
      {props.messages.map(function (message) {
        return (
          <div key={message.id} className="message-block">
            <div
              className={
                "message-bubble " + message.role + (message.error ? " error" : "")
              }
            >
              {message.content || (message.isStreaming ? "" : "…")}
            </div>
            {message.error ? (
              <ErrorActions
                error={message.error}
                disabled={props.isStreaming}
                onRetry={
                  props.onRetry
                    ? function () {
                        props.onRetry!(message.id);
                      }
                    : undefined
                }
              />
            ) : null}
            {message.role === "assistant" && message.steps && message.steps.length > 0 ? (
              <AgentTrace steps={message.steps} />
            ) : null}
            {message.role === "assistant" && message.pendingEdit ? (
              <EditPreview
                edit={message.pendingEdit}
                disabled={props.isStreaming}
                onApply={function () {
                  props.onApplyEdit(message.id);
                }}
                onReject={function () {
                  props.onRejectEdit(message.id);
                }}
                onUndo={function () {
                  props.onUndoEdit(message.id);
                }}
              />
            ) : null}
          </div>
        );
      })}
      {props.isStreaming ? (
        <div style={{ alignSelf: "flex-start", paddingLeft: 4 }}>
          <Spinner size={SpinnerSize.small} label="Thinking…" />
        </div>
      ) : null}
    </div>
  );
}