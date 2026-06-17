import { FluentProvider, Spinner, webLightTheme } from "@fluentui/react-components";
import { Suspense, lazy, useEffect, useState } from "react";
import { useChat } from "../hooks/useChat";
import { useDocumentKey } from "../hooks/useDocumentKey";
import { useSettingsStore } from "../settings/store";
import type { ContextMode } from "../types/context";
import { ChatPanel } from "./components/ChatPanel";
import { Header } from "./components/Header";
import { OnboardingWizard } from "./components/OnboardingWizard";

const SettingsPanel = lazy(() =>
  import("./components/SettingsPanel").then((module) => ({
    default: module.SettingsPanel,
  })),
);

export function App() {
  const [view, setView] = useState<"chat" | "settings">("chat");
  const [contextMode, setContextMode] = useState<ContextMode>("selection");
  const load = useSettingsStore((s) => s.load);
  const isConfigured = useSettingsStore((s) => s.isConfigured);
  const onboardingCompleted = useSettingsStore((s) => s.preferences.onboardingCompleted);
  const { docKey } = useDocumentKey();
  const chat = useChat(contextMode, docKey);

  useEffect(() => {
    load();
  }, [load]);

  useEffect(() => {
    if (!isConfigured && onboardingCompleted) {
      setView("settings");
    }
  }, [isConfigured, onboardingCompleted]);

  if (!onboardingCompleted) {
    return (
      <FluentProvider theme={webLightTheme}>
        <OnboardingWizard onComplete={() => setView(isConfigured ? "chat" : "settings")} />
      </FluentProvider>
    );
  }

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
            docKey={docKey}
            contextMode={contextMode}
            onContextModeChange={setContextMode}
            onSend={chat.sendMessage}
            onRetry={(messageId) => void chat.retryMessage(messageId)}
            onApplyEdit={(messageId) => void chat.applyEdit(messageId)}
            onRejectEdit={(messageId) => void chat.rejectEdit(messageId)}
            onUndoEdit={(messageId) => void chat.undoEdit(messageId)}
          />
        ) : (
          <div className="panel-body">
            <Suspense
              fallback={
                <div className="settings-loading">
                  <Spinner size="small" label="Loading settings…" />
                </div>
              }
            >
              <SettingsPanel />
            </Suspense>
          </div>
        )}
      </div>
    </FluentProvider>
  );
}