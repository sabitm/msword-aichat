import { Text } from "@fluentui/react";
import * as React from "react";
import type { AgentStep } from "../../types/agent";

interface AgentTraceProps {
  steps: AgentStep[];
}

function formatStep(step: AgentStep): string {
  if (step.type === "tool_call") {
    return "Called " + step.toolName;
  }
  if (step.type === "tool_result") {
    if (step.status === "pending") {
      return step.toolName + ": awaiting approval";
    }
    if (step.status === "error") {
      return step.toolName + ": failed";
    }
    return step.toolName + ": done";
  }
  return "Assistant response";
}

export function AgentTrace(props: AgentTraceProps): React.ReactElement | null {
  if (props.steps.length === 0) {
    return null;
  }

  return (
    <details className="agent-trace">
      <summary>
        <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
          Agent steps ({props.steps.length})
        </Text>
      </summary>
      <ul className="agent-trace-list">
        {props.steps.map(function (step) {
          return (
            <li key={step.id} className={"agent-trace-item status-" + step.status}>
              <span>{formatStep(step)}</span>
              {step.error ? <span className="agent-trace-error">{step.error}</span> : null}
            </li>
          );
        })}
      </ul>
    </details>
  );
}