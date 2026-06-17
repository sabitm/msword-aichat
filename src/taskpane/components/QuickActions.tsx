import { Button, Text } from "@fluentui/react-components";
import type { ContextMode } from "../../types/context";

export type QuickActionId = "summarize" | "improve" | "explain";

const QUICK_ACTIONS: Array<{
  id: QuickActionId;
  label: string;
  buildPrompt: (mode: ContextMode) => string;
}> = [
  {
    id: "summarize",
    label: "Summarize",
    buildPrompt: (mode) =>
      mode === "outline"
        ? "Summarize the document based on its outline. Highlight the main themes and structure."
        : "Summarize the document context concisely. Focus on the key points.",
  },
  {
    id: "improve",
    label: "Improve",
    buildPrompt: () =>
      "Improve the writing in the provided document context. Fix clarity, grammar, and flow. Explain what you changed.",
  },
  {
    id: "explain",
    label: "Explain",
    buildPrompt: () =>
      "Explain the provided document context clearly. Define any jargon and break down the main ideas.",
  },
];

interface QuickActionsProps {
  contextMode: ContextMode;
  disabled?: boolean;
  onAction: (prompt: string) => void;
}

export function QuickActions({ contextMode, disabled, onAction }: QuickActionsProps) {
  const needsContext = contextMode === "none";

  return (
    <div className="quick-actions">
      <Text size={200} weight="semibold">
        Quick actions
      </Text>
      <div className="quick-actions-row">
        {QUICK_ACTIONS.map((action) => (
          <Button
            key={action.id}
            appearance="secondary"
            size="small"
            disabled={disabled || needsContext}
            onClick={() => onAction(action.buildPrompt(contextMode))}
          >
            {action.label}
          </Button>
        ))}
      </div>
      {needsContext ? (
        <Text size={100} className="quick-actions-hint">
          Select Selection or Outline context to use quick actions.
        </Text>
      ) : null}
    </div>
  );
}