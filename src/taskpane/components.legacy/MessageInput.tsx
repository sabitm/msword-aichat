import { IconButton, Text, TextField } from "@fluentui/react";
import * as React from "react";
import { listMatchingSlashCommands } from "../../agent/slash-commands";

interface MessageInputProps {
  disabled?: boolean;
  onSend: (message: string) => void;
}

export function MessageInput(props: MessageInputProps): React.ReactElement {
  var _a = React.useState("");
  var value = _a[0];
  var setValue = _a[1];

  var slashMatches = listMatchingSlashCommands(value);

  function handleSend(): void {
    var trimmed = value.trim();
    if (!trimmed || props.disabled) {
      return;
    }
    props.onSend(trimmed);
    setValue("");
  }

  return (
    <div className="chat-input-shell">
      {slashMatches.length > 0 ? (
        <div className="slash-hints" role="listbox" aria-label="Slash commands">
          {slashMatches.map(function (command) {
            return (
              <button
                key={command.name}
                type="button"
                className="slash-hint-item"
                disabled={props.disabled}
                onClick={function () {
                  setValue("/" + command.name + " ");
                }}
              >
                <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                  /{command.name}
                </Text>
                <Text variant="small">{command.description}</Text>
              </button>
            );
          })}
        </div>
      ) : null}
      <div className="chat-input-bar">
        <TextField
          multiline
          autoAdjustHeight
          resizable={false}
          value={value}
          disabled={props.disabled}
          placeholder={
            props.disabled
              ? "Configure your AI endpoint in Settings"
              : "Message… (/fix, /table, /toc, /summarize, /formal, /comment)"
          }
          onChange={function (_event, nextValue) {
            setValue(nextValue || "");
          }}
          onKeyDown={function (event) {
            if (event.key === "Enter" && !event.shiftKey) {
              event.preventDefault();
              handleSend();
            }
          }}
        />
        <IconButton
          iconProps={{ iconName: "Send" }}
          title="Send message"
          ariaLabel="Send message"
          disabled={props.disabled || !value.trim()}
          onClick={handleSend}
        />
      </div>
    </div>
  );
}