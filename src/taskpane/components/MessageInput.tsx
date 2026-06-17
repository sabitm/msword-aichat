import { Button, Text, Textarea } from "@fluentui/react-components";
import { Send24Regular } from "@fluentui/react-icons";
import { useMemo, useState } from "react";
import { listMatchingSlashCommands } from "../../agent/slash-commands";

interface MessageInputProps {
  disabled?: boolean;
  onSend: (message: string) => void;
}

export function MessageInput({ disabled, onSend }: MessageInputProps) {
  const [value, setValue] = useState("");

  const slashMatches = useMemo(() => listMatchingSlashCommands(value), [value]);

  const handleSend = () => {
    const trimmed = value.trim();
    if (!trimmed || disabled) return;
    onSend(trimmed);
    setValue("");
  };

  return (
    <div className="chat-input-shell">
      {slashMatches.length > 0 ? (
        <div className="slash-hints" role="listbox" aria-label="Slash commands">
          {slashMatches.map((command) => (
            <button
              key={command.name}
              type="button"
              className="slash-hint-item"
              disabled={disabled}
              onClick={() => setValue(`/${command.name} `)}
            >
              <Text weight="semibold">/{command.name}</Text>
              <Text size={200}>{command.description}</Text>
            </button>
          ))}
        </div>
      ) : null}
      <div className="chat-input-bar">
        <Textarea
          value={value}
          disabled={disabled}
          placeholder={
            disabled
              ? "Configure your AI endpoint in Settings"
              : "Message… (/fix, /table, /toc, /summarize, /formal, /comment)"
          }
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
    </div>
  );
}