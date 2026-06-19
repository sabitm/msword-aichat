import { DefaultButton, IconButton } from "@fluentui/react";
import * as React from "react";

interface HeaderProps {
  view: "chat" | "settings";
  onViewChange: (view: "chat" | "settings") => void;
  onNewChat: () => void;
  canChat: boolean;
}

export function Header(props: HeaderProps): React.ReactElement {
  var view = props.view;
  var onViewChange = props.onViewChange;
  var onNewChat = props.onNewChat;
  var canChat = props.canChat;

  return (
    <div className="header-bar" aria-label="Word AI Chat toolbar">
      <span className="header-title">Word AI Chat</span>
      <div className="header-actions">
        <IconButton
          iconProps={{ iconName: "Chat" }}
          title="Chat"
          ariaLabel="Chat"
          checked={view === "chat"}
          styles={
            view === "chat"
              ? { root: { backgroundColor: "#deecf9" }, rootChecked: { backgroundColor: "#deecf9" } }
              : undefined
          }
          onClick={function () {
            onViewChange("chat");
          }}
        />
        <IconButton
          iconProps={{ iconName: "Settings" }}
          title="Settings"
          ariaLabel="Settings"
          checked={view === "settings"}
          styles={
            view === "settings"
              ? { root: { backgroundColor: "#deecf9" }, rootChecked: { backgroundColor: "#deecf9" } }
              : undefined
          }
          onClick={function () {
            onViewChange("settings");
          }}
        />
        <DefaultButton
          iconProps={{ iconName: "Delete" }}
          disabled={!canChat}
          onClick={onNewChat}
        >
          New chat
        </DefaultButton>
      </div>
    </div>
  );
}