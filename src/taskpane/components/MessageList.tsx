import { Spinner, Text } from "@fluentui/react-components";
import type { UiMessage } from "../../hooks/useChat";

interface MessageListProps {
  messages: UiMessage[];
  isStreaming: boolean;
}

export function MessageList({ messages, isStreaming }: MessageListProps) {
  if (messages.length === 0) {
    return (
      <div className="empty-state">
        <Text size={400} weight="semibold">
          Start a conversation
        </Text>
        <Text size={300}>
          Ask questions about your document, get writing help, or brainstorm ideas.
        </Text>
      </div>
    );
  }

  return (
    <div className="message-list">
      {messages.map((message) => (
        <div
          key={message.id}
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
      ))}
      {isStreaming ? (
        <div style={{ alignSelf: "flex-start", paddingLeft: 4 }}>
          <Spinner size="tiny" label="Thinking…" />
        </div>
      ) : null}
    </div>
  );
}