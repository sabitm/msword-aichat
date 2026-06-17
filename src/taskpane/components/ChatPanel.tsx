import { MessageBar, MessageBarBody } from "@fluentui/react-components";
import type { UiMessage } from "../../hooks/useChat";
import { useSettingsStore } from "../../settings/store";
import { MessageInput } from "./MessageInput";
import { MessageList } from "./MessageList";

interface ChatPanelProps {
  messages: UiMessage[];
  isStreaming: boolean;
  onSend: (message: string) => void;
}

export function ChatPanel({ messages, isStreaming, onSend }: ChatPanelProps) {
  const isConfigured = useSettingsStore((s) => s.isConfigured);

  return (
    <div className="chat-panel">
      <div className="panel-body">
        {!isConfigured ? (
          <MessageBar intent="warning" style={{ margin: 12 }}>
            <MessageBarBody>
              Configure your AI endpoint in Settings before chatting.
            </MessageBarBody>
          </MessageBar>
        ) : null}
        <MessageList messages={messages} isStreaming={isStreaming} />
      </div>
      <MessageInput disabled={!isConfigured || isStreaming} onSend={onSend} />
    </div>
  );
}