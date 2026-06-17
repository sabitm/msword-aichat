import {
  Button,
  makeStyles,
  shorthands,
  tokens,
  Toolbar,
  ToolbarButton,
  ToolbarDivider,
} from "@fluentui/react-components";
import {
  Chat24Regular,
  Delete24Regular,
  Settings24Regular,
} from "@fluentui/react-icons";

const useStyles = makeStyles({
  root: {
    ...shorthands.borderBottom("1px", "solid", tokens.colorNeutralStroke2),
    backgroundColor: tokens.colorNeutralBackground1,
  },
  title: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase300,
    paddingLeft: tokens.spacingHorizontalS,
  },
});

interface HeaderProps {
  view: "chat" | "settings";
  onViewChange: (view: "chat" | "settings") => void;
  onNewChat: () => void;
  canChat: boolean;
}

export function Header({ view, onViewChange, onNewChat, canChat }: HeaderProps) {
  const styles = useStyles();

  return (
    <Toolbar className={styles.root} aria-label="Word AI Chat toolbar">
      <span className={styles.title}>Word AI Chat</span>
      <ToolbarDivider />
      <ToolbarButton
        aria-label="Chat"
        icon={<Chat24Regular />}
        appearance={view === "chat" ? "primary" : "subtle"}
        onClick={() => onViewChange("chat")}
      />
      <ToolbarButton
        aria-label="Settings"
        icon={<Settings24Regular />}
        appearance={view === "settings" ? "primary" : "subtle"}
        onClick={() => onViewChange("settings")}
      />
      <ToolbarDivider />
      <Button
        appearance="subtle"
        icon={<Delete24Regular />}
        disabled={!canChat}
        onClick={onNewChat}
      >
        New chat
      </Button>
    </Toolbar>
  );
}