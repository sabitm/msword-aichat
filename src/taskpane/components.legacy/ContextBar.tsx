import { IconButton, Spinner, SpinnerSize, Text } from "@fluentui/react";
import * as React from "react";
import {
  CONTEXT_MODE_LABELS,
  type ContextMode,
  type DocumentContext,
} from "../../types/context";
import { IeSelect } from "./IeSelect";

var contextOptions = [
  { value: "selection", label: CONTEXT_MODE_LABELS.selection },
  { value: "outline", label: CONTEXT_MODE_LABELS.outline },
  { value: "none", label: CONTEXT_MODE_LABELS.none },
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
        <IeSelect
          label="Context"
          fieldClassName="context-dropdown"
          value={props.mode}
          options={contextOptions}
          disabled={props.disabled}
          onChange={function (value) {
            props.onModeChange(value as ContextMode);
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