import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { useEffect, useState } from "react";
import { useChat } from "../hooks/useChat";
import { useSettingsStore } from "../settings/store";
import type { ContextMode } from "../types/context";
import { ChatPanel } from "./components/ChatPanel";
import { Header } from "./components/Header";
import { SettingsPanel } from "./components/SettingsPanel";

export function App() {
  const [view, setView] = useState<"chat" | "settings">("chat");
  const [contextMode, setContextMode] = useState<ContextMode>("selection");
  const load = useSettingsStore((s) => s.load);
  const isConfigured = useSettingsStore((s) => s.isConfigured);
  const chat = useChat(contextMode);

  useEffect(() => {
    load();
  }, [load]);

  useEffect(() => {
    if (!isConfigured) {
      setView("settings");
    }
  }, [isConfigured]);

  return (
    <FluentProvider theme={webLightTheme}>
      <div className="app-shell">
        <Header
          view={view}
          canChat={view === "chat"}
          onViewChange={setView}
          onNewChat={chat.clearMessages}
        />
        {view === "chat" ? (
          <ChatPanel
            messages={chat.messages}
            isStreaming={chat.isStreaming}
            contextMode={contextMode}
            onContextModeChange={setContextMode}
            onSend={chat.sendMessage}
            onApplyEdit={(messageId) => void chat.applyEdit(messageId)}
            onRejectEdit={(messageId) => void chat.rejectEdit(messageId)}
            onUndoEdit={(messageId) => void chat.undoEdit(messageId)}
          />
        ) : (
          <div className="panel-body">
            <SettingsPanel />
          </div>
        )}
      </div>
    </FluentProvider>
  );
}