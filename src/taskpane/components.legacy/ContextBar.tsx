import {
  Dropdown,
  type IDropdownOption,
  IconButton,
  Spinner,
  SpinnerSize,
  Text,
} from "@fluentui/react";
import * as React from "react";
import {
  CONTEXT_MODE_LABELS,
  type ContextMode,
  type DocumentContext,
} from "../../types/context";

var contextOptions: IDropdownOption[] = [
  { key: "selection", text: CONTEXT_MODE_LABELS.selection },
  { key: "outline", text: CONTEXT_MODE_LABELS.outline },
  { key: "none", text: CONTEXT_MODE_LABELS.none },
];

interface ContextBarProps {
  mode: ContextMode;
  context: DocumentContext;
  isLoading: boolean;
  disabled?: boolean;
  onModeChange: (mode: ContextMode) => void;
  onRefresh: () => void;
}

export function ContextBar(props: ContextBarProps): React.ReactElement {
  var statusText = (function () {
    if (props.mode === "none") {
      return "No document context attached";
    }
    if (props.isLoading) {
      return "Loading context…";
    }
    if (props.context.error) {
      return props.context.error;
    }
    if (props.context.empty) {
      return props.mode === "selection"
        ? "No text selected in the document"
        : "No headings found in the document";
    }
    return (
      "~" +
      props.context.tokenEstimate.toLocaleString() +
      " tokens" +
      (props.context.truncated ? " (truncated)" : "")
    );
  })();

  return (
    <div className="context-bar">
      <div className="context-bar-row">
        <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
          Context
        </Text>
        <Dropdown
          className="context-dropdown"
          disabled={props.disabled}
          selectedKey={props.mode}
          options={contextOptions}
          onChange={function (_event, option) {
            if (option) {
              props.onModeChange(option.key as ContextMode);
            }
          }}
        />
        <IconButton
          iconProps={{ iconName: "Sync" }}
          title="Refresh document context"
          ariaLabel="Refresh context"
          disabled={props.disabled || props.mode === "none" || props.isLoading}
          onClick={props.onRefresh}
        >
          {props.isLoading ? <Spinner size={SpinnerSize.small} /> : null}
        </IconButton>
      </div>
      <Text variant="small" className="context-status">
        {statusText}
      </Text>
    </div>
  );
}