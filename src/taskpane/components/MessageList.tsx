import { Spinner, Text } from "@fluentui/react-components";
import type { UiMessage } from "../../hooks/useChat";
import { AgentTrace } from "./AgentTrace";
import { EditPreview } from "./EditPreview";

interface MessageListProps {
  messages: UiMessage[];
  isStreaming: boolean;
  onApplyEdit: (messageId: string) => void;
  onRejectEdit: (messageId: string) => void;
}

export function MessageList({
  messages,
  isStreaming,
  onApplyEdit,
  onRejectEdit,
}: MessageListProps) {
  if (messages.length === 0) {
    return (
      <div className="empty-state">
        <Text size={400} weight="semibold">
          Start a conversation
        </Text>
        <Text size={300}>
          Ask questions about your document, get writing help, or use Agent mode to edit.
        </Text>
      </div>
    );
  }

  return (
    <div className="message-list">
      {messages.map((message) => (
        <div key={message.id} className="message-block">
          <div
            className={`message-bubble ${message.role}${message.error ? " error" : ""}`}
          >
            {message.content || (message.isStreaming ? "" : "…")}
            {message.error ? (
              <div>
                <br />
                <strong>Error:</strong> {message.error}
              </div>
            ) : null}
          </div>
          {message.role === "assistant" && message.steps?.length ? (
            <AgentTrace steps={message.steps} />
          ) : null}
          {message.role === "assistant" && message.pendingEdit ? (
            <EditPreview
              edit={message.pendingEdit}
              disabled={isStreaming}
              onApply={() => onApplyEdit(message.id)}
              onReject={() => onRejectEdit(message.id)}
            />
          ) : null}
        </div>
      ))}
      {isStreaming ? (
        <div style={{ alignSelf: "flex-start", paddingLeft: 4 }}>
          <Spinner size="tiny" label="Thinking…" />
        </div>
      ) : null}
    </div>
  );
}