import { Dropdown, Option, Text } from "@fluentui/react-components";
import { useSettingsStore } from "../../settings/store";
import type { InteractionMode } from "../../types/agent";

const MODE_LABELS: Record<InteractionMode, string> = {
  chat: "Chat",
  agent: "Agent",
};

export function ModeBar() {
  const preferences = useSettingsStore((s) => s.preferences);
  const setInteractionMode = useSettingsStore((s) => s.setInteractionMode);
  const save = useSettingsStore((s) => s.save);

  return (
    <div className="mode-bar">
      <Text size={200} weight="semibold">
        Mode
      </Text>
      <Dropdown
        className="mode-dropdown"
        value={MODE_LABELS[preferences.interactionMode]}
        selectedOptions={[preferences.interactionMode]}
        onOptionSelect={(_event, data) => {
          const mode = data.optionValue as InteractionMode | undefined;
          if (!mode) return;
          setInteractionMode(mode);
          save();
        }}
      >
        <Option value="chat">{MODE_LABELS.chat}</Option>
        <Option value="agent">{MODE_LABELS.agent}</Option>
      </Dropdown>
      <Text size={100} className="mode-hint">
        {preferences.interactionMode === "agent"
          ? "Agent can call document tools. Edits require approval unless auto-apply is enabled."
          : "Chat streams responses without document tools."}
      </Text>
    </div>
  );
}