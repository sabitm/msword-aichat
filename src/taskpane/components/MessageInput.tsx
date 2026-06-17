import { Button, Textarea } from "@fluentui/react-components";
import { Send24Regular } from "@fluentui/react-icons";
import { useState } from "react";

interface MessageInputProps {
  disabled?: boolean;
  onSend: (message: string) => void;
}

export function MessageInput({ disabled, onSend }: MessageInputProps) {
  const [value, setValue] = useState("");

  const handleSend = () => {
    const trimmed = value.trim();
    if (!trimmed || disabled) return;
    onSend(trimmed);
    setValue("");
  };

  return (
    <div className="chat-input-bar">
      <Textarea
        value={value}
        disabled={disabled}
        placeholder={disabled ? "Configure your AI endpoint in Settings" : "Message…"}
        onChange={(_event, data) => setValue(data.value)}
        onKeyDown={(event) => {
          if (event.key === "Enter" && !event.shiftKey) {
            event.preventDefault();
            handleSend();
          }
        }}
      />
      <Button
        appearance="primary"
        icon={<Send24Regular />}
        disabled={disabled || !value.trim()}
        onClick={handleSend}
      >
        Send
      </Button>
    </div>
  );
}