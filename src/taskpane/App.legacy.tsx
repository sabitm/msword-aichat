import * as React from "react";
import { useSettingsStore } from "../hooks/useSettingsStore.legacy";
import { settingsStore } from "../settings/store.legacy";
import { ChatPlaceholder } from "./components.legacy/ChatPlaceholder";
import { Header } from "./components.legacy/Header";
import { SettingsPanel } from "./components.legacy/SettingsPanel";

interface AppLegacyProps {
  hostLabel: string;
}

export function AppLegacy(_props: AppLegacyProps): React.ReactElement {
  var _a = React.useState<"chat" | "settings">("chat");
  var view = _a[0];
  var setView = _a[1];

  var storeState = useSettingsStore();
  var isConfigured = storeState.isConfigured;

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

  function handleNewChat(): void {
    // Chat persistence arrives in IE-2 / IE-4.
  }

  return (
    <div className="app-shell">
      <Header
        view={view}
        canChat={view === "chat" && isConfigured}
        onViewChange={setView}
        onNewChat={handleNewChat}
      />
      {view === "chat" ? (
        <div className="panel-body">
          <ChatPlaceholder isConfigured={isConfigured} />
        </div>
      ) : (
        <div className="panel-body">
          <SettingsPanel />
        </div>
      )}
    </div>
  );
}