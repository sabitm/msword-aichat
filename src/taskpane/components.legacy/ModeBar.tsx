import { Dropdown, type IDropdownOption, Text } from "@fluentui/react";
import * as React from "react";
import { useSettingsStore } from "../../hooks/useSettingsStore.legacy";
import { settingsStore } from "../../settings/store.legacy";
import type { InteractionMode } from "../../types/agent";

var MODE_LABELS: Record<InteractionMode, string> = {
  chat: "Chat",
  agent: "Agent",
};

var modeOptions: IDropdownOption[] = [
  { key: "chat", text: MODE_LABELS.chat },
  { key: "agent", text: MODE_LABELS.agent },
];

export function ModeBar(): React.ReactElement {
  var preferences = useSettingsStore().preferences;

  return (
    <div className="mode-bar">
      <Text variant="small" block>
        Mode
      </Text>
      <Dropdown
        className="mode-dropdown"
        selectedKey={preferences.interactionMode}
        options={modeOptions}
        onChange={function (_event, option) {
          if (!option) {
            return;
          }
          settingsStore.setInteractionMode(option.key as InteractionMode);
          settingsStore.save();
        }}
      />
      <Text variant="small" block className="mode-hint">
        {preferences.interactionMode === "agent"
          ? "Agent tools arrive in IE-3. Use Chat mode to stream responses now."
          : "Chat streams responses without document tools."}
      </Text>
    </div>
  );
}