import {
  CommandBar,
  type ICommandBarItemProps,
  mergeStyles,
} from "@fluentui/react";
import * as React from "react";

var headerRootClass = mergeStyles({
  borderBottom: "1px solid #e0e0e0",
  backgroundColor: "#fff",
});

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

  var items: ICommandBarItemProps[] = [
    {
      key: "title",
      text: "Word AI Chat",
      disabled: true,
      styles: { root: { fontWeight: 600 } },
    },
    {
      key: "chat",
      text: "Chat",
      iconProps: { iconName: "Chat" },
      buttonStyles: view === "chat" ? { root: { backgroundColor: "#deecf9" } } : undefined,
      onClick: function () {
        onViewChange("chat");
      },
    },
    {
      key: "settings",
      text: "Settings",
      iconProps: { iconName: "Settings" },
      buttonStyles: view === "settings" ? { root: { backgroundColor: "#deecf9" } } : undefined,
      onClick: function () {
        onViewChange("settings");
      },
    },
    {
      key: "newChat",
      text: "New chat",
      iconProps: { iconName: "Delete" },
      disabled: !canChat,
      onClick: onNewChat,
    },
  ];

  return (
    <div className={headerRootClass}>
      <CommandBar items={items} ariaLabel="Word AI Chat toolbar" />
    </div>
  );
}