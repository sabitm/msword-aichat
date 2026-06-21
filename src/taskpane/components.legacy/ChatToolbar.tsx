import { IconButton, Spinner, SpinnerSize } from "@fluentui/react";
import * as React from "react";
import { useSettingsStore } from "../../hooks/useSettingsStore.legacy";
import { settingsStore } from "../../settings/store.legacy";
import type { InteractionMode } from "../../types/agent";
import {
  CONTEXT_MODE_LABELS,
  type ContextMode,
  type DocumentContext,
} from "../../types/context";
import { formatDocumentKeyLabel } from "../../word/document-key.legacy";
import { IeSelect } from "./IeSelect";

var MODE_LABELS: Record<InteractionMode, string> = {
  chat: "Chat",
  agent: "Agent",
};

var modeOptions = [
  { value: "chat", label: MODE_LABELS.chat },
  { value: "agent", label: MODE_LABELS.agent },
];

var contextOptions = [
  { value: "selection", label: CONTEXT_MODE_LABELS.selection },
  { value: "outline", label: CONTEXT_MODE_LABELS.outline },
  { value: "none", label: CONTEXT_MODE_LABELS.none },
];

interface ChatToolbarProps {
  contextMode: ContextMode;
  context: DocumentContext;
  isLoading: boolean;
  docKey: string;
  persistConversations: boolean;
  disabled?: boolean;
  onContextModeChange: (mode: ContextMode) => void;
  onRefresh: () => void;
}

function formatContextStatus(mode: ContextMode, context: DocumentContext, isLoading: boolean): string {
  if (mode === "none") {
    return "No document context";
  }
  if (isLoading) {
    return "Loading context…";
  }
  if (context.error) {
    return context.error;
  }
  if (context.empty) {
    return mode === "selection" ? "No text selected" : "No headings found";
  }
  var tokenLabel = "~" + context.tokenEstimate.toLocaleString() + " tokens";
  if (context.truncated) {
    tokenLabel += " (truncated)";
  }
  if (mode === "selection" && context.tableSelection) {
    tokenLabel +=
      " · table " +
      (context.tableSelection.tableIndex + 1) +
      ", row " +
      (context.tableSelection.rowIndex + 1);
  }
  return tokenLabel;
}

export function ChatToolbar(props: ChatToolbarProps): React.ReactElement {
  var preferences = useSettingsStore().preferences;
  var interactionMode = preferences.interactionMode;

  var _a = React.useState(false);
  var expanded = _a[0];
  var setExpanded = _a[1];

  var contextStatus = formatContextStatus(props.contextMode, props.context, props.isLoading);

  var statusParts: string[] = [];
  if (
    props.persistConversations &&
    props.docKey !== "loading" &&
    props.docKey !== "browser"
  ) {
    statusParts.push("Saved: " + formatDocumentKeyLabel(props.docKey));
  }
  statusParts.push(contextStatus);
  var statusLine = statusParts.join(" · ");

  var modeHint =
    interactionMode === "agent"
      ? "Agent can call document tools. Edits require approval unless auto-apply is enabled."
      : "Chat streams responses without document tools.";

  return (
    <div className="chat-toolbar">
      <div className="chat-toolbar-row">
        <IeSelect
          ariaLabel="Interaction mode"
          className="chat-toolbar-select"
          fieldClassName="chat-toolbar-field"
          value={interactionMode}
          options={modeOptions}
          onChange={function (value) {
            settingsStore.setInteractionMode(value as InteractionMode);
            settingsStore.save();
          }}
        />
        <IeSelect
          ariaLabel="Document context"
          className="chat-toolbar-select"
          fieldClassName="chat-toolbar-field"
          value={props.contextMode}
          options={contextOptions}
          disabled={props.disabled}
          onChange={function (value) {
            props.onContextModeChange(value as ContextMode);
          }}
        />
        <IconButton
          className="chat-toolbar-refresh"
          iconProps={{ iconName: "Sync" }}
          title="Refresh document context"
          ariaLabel="Refresh context"
          disabled={props.disabled || props.contextMode === "none" || props.isLoading}
          onClick={props.onRefresh}
        >
          {props.isLoading ? <Spinner size={SpinnerSize.small} /> : null}
        </IconButton>
        <IconButton
          className="chat-toolbar-toggle"
          iconProps={{ iconName: expanded ? "ChevronUp" : "ChevronDown" }}
          title={expanded ? "Hide options" : "Show options"}
          ariaLabel={expanded ? "Hide chat options" : "Show chat options"}
          onClick={function () {
            setExpanded(!expanded);
          }}
        />
      </div>
      <div className="chat-toolbar-status" title={statusLine}>
        {statusLine}
      </div>
      {expanded ? (
        <div className="chat-toolbar-details">
          <p className="chat-toolbar-hint">{modeHint}</p>
        </div>
      ) : null}
    </div>
  );
}