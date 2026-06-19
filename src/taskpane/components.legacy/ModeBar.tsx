import { Text } from "@fluentui/react";
import * as React from "react";
import { useSettingsStore } from "../../hooks/useSettingsStore.legacy";
import { settingsStore } from "../../settings/store.legacy";
import type { InteractionMode } from "../../types/agent";
import { IeSelect } from "./IeSelect";

var MODE_LABELS: Record<InteractionMode, string> = {
  chat: "Chat",
  agent: "Agent",
};

var modeOptions = [
  { value: "chat", label: MODE_LABELS.chat },
  { value: "agent", label: MODE_LABELS.agent },
];

export function ModeBar(): React.ReactElement {
  var preferences = useSettingsStore().preferences;

  return (
    <div className="mode-bar">
      <IeSelect
        label="Mode"
        fieldClassName="mode-dropdown"
        value={preferences.interactionMode}
        options={modeOptions}
        onChange={function (value) {
          settingsStore.setInteractionMode(value as InteractionMode);
          settingsStore.save();
        }}
      />
      <Text variant="small" block className="mode-hint">
        {preferences.interactionMode === "agent"
          ? "Agent can call document tools. Edits require approval unless auto-apply is enabled."
          : "Chat streams responses without document tools."}
      </Text>
    </div>
  );
}