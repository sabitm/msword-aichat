import { Text } from "@fluentui/react-components";
import type { AgentStep } from "../../types/agent";

interface AgentTraceProps {
  steps: AgentStep[];
}

function formatStep(step: AgentStep): string {
  if (step.type === "tool_call") {
    return `Called ${step.toolName}`;
  }
  if (step.type === "tool_result") {
    if (step.status === "pending") {
      return `${step.toolName}: awaiting approval`;
    }
    if (step.status === "error") {
      return `${step.toolName}: failed`;
    }
    return `${step.toolName}: done`;
  }
  return "Assistant response";
}

export function AgentTrace({ steps }: AgentTraceProps) {
  if (steps.length === 0) return null;

  return (
    <details className="agent-trace">
      <summary>
        <Text size={200} weight="semibold">
          Agent steps ({steps.length})
        </Text>
      </summary>
      <ul className="agent-trace-list">
        {steps.map((step) => (
          <li key={step.id} className={`agent-trace-item status-${step.status}`}>
            <span>{formatStep(step)}</span>
            {step.error ? <span className="agent-trace-error">{step.error}</span> : null}
          </li>
        ))}
      </ul>
    </details>
  );
}