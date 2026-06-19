import * as React from "react";
import { useChat } from "../hooks/useChat.legacy";
import { useDocumentKey } from "../hooks/useDocumentKey.legacy";
import { useSettingsStore } from "../hooks/useSettingsStore.legacy";
import { settingsStore } from "../settings/store.legacy";
import type { ContextMode } from "../types/context";
import { ChatPanel } from "./components.legacy/ChatPanel";
import { Header } from "./components.legacy/Header";
import { SettingsPanel } from "./components.legacy/SettingsPanel";

interface AppLegacyProps {
  hostLabel: string;
}

export function AppLegacy(_props: AppLegacyProps): React.ReactElement {
  var _a = React.useState<"chat" | "settings">("chat");
  var view = _a[0];
  var setView = _a[1];

  var _b = React.useState<ContextMode>("selection");
  var contextMode = _b[0];
  var setContextMode = _b[1];

  var storeState = useSettingsStore();
  var isConfigured = storeState.isConfigured;

  var documentKeyState = useDocumentKey();
  var docKey = documentKeyState.docKey;

  var chat = useChat(contextMode, docKey);

  React.useEffect(function () {
    settingsStore.load();
  }, []);

  React.useEffect(
    function () {
      if (!isConfigured) {
        setView("settings");
      }
    },
    [isConfigured],
  );

  return (
    <div className="app-shell">
      <Header
        view={view}
        canChat={view === "chat" && isConfigured}
        onViewChange={setView}
        onNewChat={chat.clearMessages}
      />
      {view === "chat" ? (
        <ChatPanel
          messages={chat.messages}
          isStreaming={chat.isStreaming}
          docKey={docKey}
          contextMode={contextMode}
          onContextModeChange={setContextMode}
          onSend={chat.sendMessage}
          onRetry={chat.retryMessage}
        />
      ) : (
        <div className="panel-body">
          <SettingsPanel />
        </div>
      )}
    </div>
  );
}