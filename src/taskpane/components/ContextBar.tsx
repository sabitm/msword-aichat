import {
  Button,
  Dropdown,
  Option,
  Spinner,
  Text,
  Tooltip,
} from "@fluentui/react-components";
import { ArrowClockwise24Regular } from "@fluentui/react-icons";
import type { DocumentContext } from "../../types/context";
import {
  CONTEXT_MODE_LABELS,
  type ContextMode,
} from "../../types/context";

interface ContextBarProps {
  mode: ContextMode;
  context: DocumentContext;
  isLoading: boolean;
  disabled?: boolean;
  onModeChange: (mode: ContextMode) => void;
  onRefresh: () => void;
}

export function ContextBar({
  mode,
  context,
  isLoading,
  disabled,
  onModeChange,
  onRefresh,
}: ContextBarProps) {
  const statusText = (() => {
    if (mode === "none") return "No document context attached";
    if (isLoading) return "Loading context…";
    if (context.error) return context.error;
    if (context.empty) {
      return mode === "selection"
        ? "No text selected in the document"
        : "No headings found in the document";
    }
    return `~${context.tokenEstimate.toLocaleString()} tokens${
      context.truncated ? " (truncated)" : ""
    }`;
  })();

  return (
    <div className="context-bar">
      <div className="context-bar-row">
        <Text size={200} weight="semibold">
          Context
        </Text>
        <Dropdown
          className="context-dropdown"
          disabled={disabled}
          value={CONTEXT_MODE_LABELS[mode]}
          selectedOptions={[mode]}
          onOptionSelect={(_event, data) => {
            const next = data.optionValue as ContextMode | undefined;
            if (next) onModeChange(next);
          }}
        >
          <Option value="selection">{CONTEXT_MODE_LABELS.selection}</Option>
          <Option value="outline">{CONTEXT_MODE_LABELS.outline}</Option>
          <Option value="none">{CONTEXT_MODE_LABELS.none}</Option>
        </Dropdown>
        <Tooltip content="Refresh document context" relationship="label">
          <Button
            appearance="subtle"
            icon={isLoading ? <Spinner size="tiny" /> : <ArrowClockwise24Regular />}
            disabled={disabled || mode === "none" || isLoading}
            onClick={onRefresh}
            aria-label="Refresh context"
          />
        </Tooltip>
      </div>
      <Text size={200} className="context-status">
        {statusText}
      </Text>
    </div>
  );
}