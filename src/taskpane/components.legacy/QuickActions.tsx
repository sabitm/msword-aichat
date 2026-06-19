import { DefaultButton, Text } from "@fluentui/react";
import * as React from "react";
import type { ContextMode } from "../../types/context";

export type QuickActionId = "summarize" | "improve" | "explain";

var QUICK_ACTIONS: Array<{
  id: QuickActionId;
  label: string;
  buildPrompt: (mode: ContextMode) => string;
}> = [
  {
    id: "summarize",
    label: "Summarize",
    buildPrompt: function (mode) {
      return mode === "outline"
        ? "Summarize the document based on its outline. Highlight the main themes and structure."
        : "Summarize the document context concisely. Focus on the key points.";
    },
  },
  {
    id: "improve",
    label: "Improve",
    buildPrompt: function () {
      return "Improve the writing in the provided document context. Fix clarity, grammar, and flow. Explain what you changed.";
    },
  },
  {
    id: "explain",
    label: "Explain",
    buildPrompt: function () {
      return "Explain the provided document context clearly. Define any jargon and break down the main ideas.";
    },
  },
];

interface QuickActionsProps {
  contextMode: ContextMode;
  disabled?: boolean;
  onAction: (prompt: string) => void;
}

export function QuickActions(props: QuickActionsProps): React.ReactElement {
  var needsContext = props.contextMode === "none";

  return (
    <div className="quick-actions">
      <Text variant="small" block styles={{ root: { fontWeight: 600 } }}>
        Quick actions
      </Text>
      <div className="quick-actions-row">
        {QUICK_ACTIONS.map(function (action) {
          return (
            <DefaultButton
              key={action.id}
              disabled={props.disabled || needsContext}
              onClick={function () {
                props.onAction(action.buildPrompt(props.contextMode));
              }}
            >
              {action.label}
            </DefaultButton>
          );
        })}
      </div>
      {needsContext ? (
        <Text variant="small" className="quick-actions-hint">
          Select Selection or Outline context to use quick actions.
        </Text>
      ) : null}
    </div>
  );
}