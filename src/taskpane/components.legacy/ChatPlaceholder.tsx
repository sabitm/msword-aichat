import { MessageBar, MessageBarType, Text } from "@fluentui/react";
import * as React from "react";

interface ChatPlaceholderProps {
  isConfigured: boolean;
}

export function ChatPlaceholder(props: ChatPlaceholderProps): React.ReactElement {
  if (!props.isConfigured) {
    return (
      <div className="empty-state">
        <Text variant="large" block>
          Configure your AI provider
        </Text>
        <Text variant="medium" block>
          Open Settings, enter your base URL, API key, and model, then save.
        </Text>
        <MessageBar messageBarType={MessageBarType.warning} isMultiline>
          Chat is disabled until settings are saved.
        </MessageBar>
      </div>
    );
  }

  return (
    <div className="empty-state">
      <Text variant="large" block>
        Chat coming in IE-2
      </Text>
      <Text variant="medium" block>
        Settings are saved. Message list, streaming, and document context will arrive in the next
        phase.
      </Text>
      <MessageBar messageBarType={MessageBarType.info} isMultiline>
        Use Settings to test your connection or fetch models while we finish the chat UI.
      </MessageBar>
    </div>
  );
}