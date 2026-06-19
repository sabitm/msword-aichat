import { Spinner, SpinnerSize, Text } from "@fluentui/react";
import * as React from "react";
import type { UiMessage } from "../../hooks/useChat.legacy";
import { ErrorActions } from "./ErrorActions";

interface MessageListProps {
  messages: UiMessage[];
  isStreaming: boolean;
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
          Ask questions about your document or use quick actions with selection or outline context.
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