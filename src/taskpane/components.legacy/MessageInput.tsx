import { IconButton, TextField } from "@fluentui/react";
import * as React from "react";

interface MessageInputProps {
  disabled?: boolean;
  onSend: (message: string) => void;
}

export function MessageInput(props: MessageInputProps): React.ReactElement {
  var _a = React.useState("");
  var value = _a[0];
  var setValue = _a[1];

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
              : "Message the assistant about your document…"
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